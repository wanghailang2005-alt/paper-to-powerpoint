"""Insert and convert SVG pages using real Microsoft PowerPoint on Windows."""
from pathlib import Path
import argparse, base64, ctypes, hashlib, json, re, time, zipfile
import xml.etree.ElementTree as ET

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def leaves(shapes):
    result = []
    for i in range(1, shapes.Count + 1):
        sh = shapes(i)
        result.extend(leaves(sh.GroupItems) if sh.Type == 6 else [sh])
    return result

def bounds(shape):
    tr = shape.TextFrame2.TextRange
    return [float(tr.BoundLeft), float(tr.BoundTop), float(tr.BoundWidth), float(tr.BoundHeight)]

def svg_info(path):
    root = ET.parse(path).getroot()
    ns = {'s': 'http://www.w3.org/2000/svg'}
    box = [float(v) for v in root.attrib['viewBox'].replace(',', ' ').split()]
    if len(box) != 4 or box[:2] != [0, 0] or min(box[2:]) <= 0:
        raise ValueError('SVG must use a positive canvas starting at 0,0')
    text = [''.join(t.itertext()) for t in root.findall('.//s:text', ns)]
    images = root.findall('.//s:image', ns)
    image_hashes = []
    for im in images:
        href = im.get('{http://www.w3.org/1999/xlink}href', im.get('href', ''))
        if not href.startswith(('data:image/png;base64,', 'data:image/jpeg;base64,')):
            raise ValueError('SVG figures must be complete embedded PNG/JPEG images')
        image_hashes.append(hashlib.sha256(base64.b64decode(href.split(',', 1)[1], validate=True)).hexdigest())
    return {'file': path.name, 'sha256': sha(path), 'canvas': box[2:], 'texts': text,
            'images': len(images), 'image_hashes': image_hashes}

def inspect(slide, expected, width, height):
    shapes = leaves(slide.Shapes)
    texts = [s for s in shapes if s.HasTextFrame and s.TextFrame.HasText]
    actual = [s.TextFrame.TextRange.Text for s in texts]
    if actual != expected['texts']:
        raise AssertionError(f'Text conversion mismatch on {expected["file"]}: {actual!r}')
    pictures = [s for s in shapes if s.Type in [11, 13]]
    if len(pictures) != expected['images']:
        raise AssertionError('Picture count changed during conversion')
    if any(s.Type in [28, 29] for s in shapes):
        raise AssertionError('Unconverted SVG remains')
    boxes = [{'text': s.TextFrame.TextRange.Text, 'bbox_pt': bounds(s),
              'size_pt': float(s.TextFrame.TextRange.Font.Size)} for s in texts]
    overflow = [b for b in boxes if b['bbox_pt'][0] < -.5 or b['bbox_pt'][1] < -.5
                or b['bbox_pt'][0] + b['bbox_pt'][2] > width + .5
                or b['bbox_pt'][1] + b['bbox_pt'][3] > height + .5]
    return texts, {'text_boxes': len(texts), 'pictures': len(pictures),
                   'other_shapes': len(shapes) - len(texts) - len(pictures),
                   'text_content_verified': True, 'text_bounds': boxes, 'overflow': overflow}

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--svg-dir', type=Path, required=True)
    ap.add_argument('--out', type=Path, required=True)
    ap.add_argument('--template', type=Path)
    ap.add_argument('--size', type=float, nargs=2, default=[960, 540], metavar=('WIDTH_PT', 'HEIGHT_PT'))
    ap.add_argument('--notes', type=Path)
    ap.add_argument('--adjustments', type=Path)
    ap.add_argument('--validate-only', action='store_true')
    ap.add_argument('--overwrite', action='store_true')
    args = ap.parse_args()
    if args.template and args.template.suffix.lower() != '.pptx':
        ap.error('Use a .pptx template')
    files = sorted(args.svg_dir.resolve().glob('slide-*.svg'), key=lambda p: int(p.stem.split('-')[-1]))
    if not files or [int(p.stem.split('-')[-1]) for p in files] != list(range(1, len(files) + 1)):
        ap.error('SVG pages must be numbered slide-01.svg through slide-NN.svg without gaps')
    expected = [svg_info(p) for p in files]
    notes = json.loads(args.notes.read_text(encoding='utf-8-sig')) if args.notes else [{} for _ in files]
    if len(notes) != len(files):
        ap.error('Notes count must match SVG pages')
    config = json.loads(args.adjustments.read_text(encoding='utf-8-sig')) if args.adjustments else {}
    if set(config) - {'title_pt', 'source_min_pt', 'page_number_pt'}:
        ap.error('Unknown adjustment key')
    if any(not isinstance(v, (float, int)) or v <= 0 for v in config.values()):
        ap.error('Adjustment sizes must be positive numbers')
    out = args.out.resolve()
    if out.exists() and any(out.iterdir()) and not args.overwrite:
        ap.error('Output directory is not empty; use a new directory or --overwrite')
    for name in ['validation/inserted', 'validation/converted', 'validation/edit_view', 'previews']:
        (out / name).mkdir(parents=True, exist_ok=True)
    template = args.template.resolve() if args.template else None
    template_hash = sha(template) if template else None
    pptx = out / 'presentation.pptx'
    if template and template == pptx:
        ap.error('Output cannot replace the original template')
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except (AttributeError, OSError):
        pass
    import win32com.client, win32gui
    from PIL import ImageGrab, Image, ImageChops, ImageStat
    app = win32com.client.Dispatch('PowerPoint.Application')
    app.Visible = True
    report = {'application': 'Microsoft PowerPoint', 'version': app.Version,
              'status': 'running', 'pages': [], 'adjustments': [], 'inputs': expected,
              'visual_review': 'Required: compare all two-state/final PNGs; pixel differences are not font-metric guarantees.'}
    p = saved = None
    def show(slide):
        app.ActiveWindow.View.GotoSlide(slide.SlideIndex)
        app.ActiveWindow.Selection.Unselect()
        time.sleep(.12)
    def export(slide, path):
        show(slide)
        slide.Export(str(path), 'PNG', 1280, round(1280 * height / width))
    def edit_capture(slide, n):
        show(slide)
        handle = win32gui.FindWindow('PPTFrameClass', None)
        if not handle:
            raise RuntimeError('PowerPoint editing window was not found')
        # ImageGrab captures visible screen pixels, including occluding windows.
        # Activate PowerPoint and verify focus instead of labeling another app as evidence.
        import win32api, win32con
        win32gui.ShowWindow(handle, win32con.SW_RESTORE)
        win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
        win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32gui.SetForegroundWindow(handle)
        time.sleep(.35)
        if win32gui.GetForegroundWindow() != handle:
            raise RuntimeError('PowerPoint could not be brought to foreground; editing screenshot was not captured')
        ImageGrab.grab(bbox=win32gui.GetWindowRect(handle)).save(out / 'validation/edit_view' / f'slide-{n:02}.png')
    try:
        # Untitled=True opens a separate copy even if the user already has the template open.
        p = app.Presentations.Open(str(template), True, True, True) if template else app.Presentations.Add()
        if not template:
            if min(args.size) <= 0:
                raise ValueError('Slide dimensions must be positive')
            p.PageSetup.SlideWidth, p.PageSetup.SlideHeight = args.size
        width, height = float(p.PageSetup.SlideWidth), float(p.PageSetup.SlideHeight)
        report['page_size_pt'] = [width, height]
        report['insertion_pt'] = [0, 0, width, height]
        for item in expected:
            if abs(item['canvas'][0] / item['canvas'][1] - width / height) > .001:
                raise ValueError('SVG and actual template aspect ratios differ; re-layout instead of stretching')
        while p.Slides.Count:
            p.Slides(1).Delete()
        app.ActiveWindow.ViewType = 9
        for n, (file, source, note) in enumerate(zip(files, expected, notes), 1):
            inserted = p.Slides.Add(n, 12)
            inserted.FollowMasterBackground = False
            inserted.Background.Fill.Solid()
            inserted.Background.Fill.ForeColor.RGB = 16777215
            inserted.DisplayMasterShapes = False
            inserted.Shapes.AddPicture(str(file), False, True, 0, 0, width, height)
            before = out / 'validation/inserted' / f'slide-{n:02}.png'
            after = out / 'validation/converted' / f'slide-{n:02}.png'
            export(inserted, before)
            slide = inserted.Duplicate().Item(1)
            app.ActiveWindow.View.GotoSlide(slide.SlideIndex)
            slide.Shapes(1).Select()
            if not app.CommandBars.GetEnabledMso('SVGEdit'):
                raise RuntimeError('This Microsoft PowerPoint installation cannot execute Convert to Shape (SVGEdit)')
            app.CommandBars.ExecuteMso('SVGEdit')
            export(slide, after)
            inserted.Delete()
            texts, stats = inspect(slide, source, width, height)
            with Image.open(before) as a, Image.open(after) as b:
                stats['two_state_mean_absolute_channel_difference'] = ImageStat.Stat(ImageChops.difference(a.convert('RGB'), b.convert('RGB'))).mean
            for i, sh in enumerate(texts, 1):
                tr = sh.TextFrame.TextRange
                old = bounds(sh)
                size = float(tr.Font.Size)
                target, role = None, None
                # Geometric defaults apply only when the caller supplies adjustment sizes.
                if old[1] < height * .14 and i == 1:
                    target, role = config.get('title_pt'), 'title'
                elif re.fullmatch(r'\d+\s*/\s*\d+', tr.Text) and old[1] > height * .85:
                    target, role = config.get('page_number_pt'), 'folio'
                elif old[1] >= height * .89:
                    target, role = max(size, config.get('source_min_pt', size)), 'source'
                sh.Name = f'Text_{i:03}'
                sh.TextFrame.WordWrap = False
                sh.TextFrame.AutoSize = 0
                if target is not None and abs(target - size) > .01:
                    tr.Font.Size = target
                    if role == 'folio':
                        new = bounds(sh)
                        sh.Left += old[0] + old[2] - new[0] - new[2]
                    report['adjustments'].append({'page': n, 'role': role, 'text': tr.Text,
                        'old_size_pt': size, 'new_size_pt': target, 'before': old, 'after': bounds(sh)})
            for holder in slide.NotesPage.Shapes:
                if holder.Type == 14 and holder.PlaceholderFormat.Type == 2:
                    holder.TextFrame.TextRange.Text = note.get('note', '')
                    break
            _, final_stats = inspect(slide, source, width, height)
            stats['final'] = final_stats
            report['pages'].append({'page': n, **stats})
            if final_stats['overflow']:
                raise AssertionError(f'Visible text exceeds slide boundary on page {n}: {final_stats["overflow"]}')
            export(slide, out / 'previews' / f'slide-{n:02}.png')
            edit_capture(slide, n)
            print(f'PowerPoint converted and checked {n}/{len(files)}', flush=True)
        if not args.validate_only:
            p.SaveAs(str(pptx), 24)
            report['pptx'] = str(pptx)
            try:
                pdf = out / 'preview.pdf'
                p.SaveAs(str(pdf), 32)
                report['pdf'] = str(pdf)
            except Exception as exc:
                report['pdf_error'] = str(exc)
            p.Saved = True
            p.Close()
            p = None
            saved = app.Presentations.Open(str(pptx), True, False, True)
            app.ActiveWindow.ViewType = 9
            if saved.Slides.Count != len(files):
                raise AssertionError('Saved slide count changed')
            report['reopened_pages'] = []
            for n, source in enumerate(expected, 1):
                slide = saved.Slides(n)
                _, stats = inspect(slide, source, width, height)
                if stats['text_bounds'] != report['pages'][n-1]['final']['text_bounds']:
                    raise AssertionError('Saved font sizes or text bounds changed')
                note_text = ''
                for holder in slide.NotesPage.Shapes:
                    if holder.Type == 14 and holder.PlaceholderFormat.Type == 2:
                        note_text = holder.TextFrame.TextRange.Text
                        break
                if note_text != notes[n-1].get('note', ''):
                    raise AssertionError('Saved presenter notes changed')
                stats['notes_verified'] = True
                report['reopened_pages'].append(stats)
                export(slide, out / 'previews' / f'slide-{n:02}.png')
                edit_capture(slide, n)
            report['reopened_in_powerpoint'] = True
            with zipfile.ZipFile(pptx) as z:
                if z.testzip() is not None:
                    raise AssertionError('PPTX zip corruption')
                report['pptx_svg_media_count'] = sum(x.lower().endswith('.svg') for x in z.namelist())
                if report['pptx_svg_media_count']:
                    raise AssertionError('SVG media remains in saved deck')
                media_hashes = {hashlib.sha256(z.read(x)).hexdigest() for x in z.namelist() if x.startswith('ppt/media/')}
                report['figure_bytes_preserved'] = all(h in media_hashes for item in expected for h in item['image_hashes'])
                if not report['figure_bytes_preserved']:
                    raise AssertionError('Original figure bytes were changed during PowerPoint save')
        report['status'] = 'structural_checks_passed_visual_review_required'
    except Exception as exc:
        report['status'] = 'failed'
        report['error'] = repr(exc)
        raise
    finally:
        for owned in [saved, p]:
            if owned is not None:
                try:
                    owned.Saved = True
                    owned.Close()
                except Exception as exc:
                    report.setdefault('close_errors', []).append(str(exc))
        if template:
            report['template_unchanged'] = sha(template) == template_hash
            if not report['template_unchanged']:
                report['status'] = 'failed'
                report['error'] = 'Original template hash changed'
        (out / 'validation/powerpoint_report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'Completed: {out}. Review final and two-state previews before delivery.', flush=True)

if __name__ == '__main__':
    main()
