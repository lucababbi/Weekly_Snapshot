from pathlib import Path
import pymupdf
import tempfile
import subprocess
from PIL import Image, ImageChops


def Escape_AppleScript_Text(value: str) -> str:
    return (
        value.replace("\\", "\\\\")
             .replace('"', '\\"')
             .replace("\n", "\\n")
    )


def autocrop_white_borders(image_path: str, padding: int = 10) -> str:
    img = Image.open(image_path).convert("RGB")
    bg = Image.new("RGB", img.size, (255, 255, 255))
    diff = ImageChops.difference(img, bg)
    bbox = diff.getbbox()

    if bbox:
        left, top, right, bottom = bbox
        left = max(0, left - padding)
        top = max(0, top - padding)
        right = min(img.width, right + padding)
        bottom = min(img.height, bottom + padding)
        img = img.crop((left, top, right, bottom))
        img.save(image_path)

    return image_path


def pdf_to_images(pdf_path: str, dpi: int = 200, max_pages: int = 5) -> list[str]:
    pdf_file = Path(pdf_path).expanduser().resolve()
    if not pdf_file.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_file}")

    temp_dir = Path(tempfile.mkdtemp(prefix="outlook_pdf_"))
    img_paths = []

    with pymupdf.open(pdf_file) as doc:
        for page_num in range(min(max_pages, len(doc))):
            page = doc[page_num]
            pix = page.get_pixmap(matrix=pymupdf.Matrix(dpi / 72, dpi / 72))
            img_path = temp_dir / f"page_{page_num + 1}.png"
            pix.save(str(img_path))
            autocrop_white_borders(str(img_path), padding=10)
            img_paths.append(str(img_path))

    return img_paths


def OutlookEmail(
    pdf_path: str,
    to_emails: list[str] | None = None,
    cc_emails: list[str] | None = None,
    subject: str = "PDF screenshots",
    body_text: str = "",
    dpi: int = 200,
    max_pages: int = 1,
    send_automatically: bool = False,
):
    if to_emails is None:
        to_emails = []

    if cc_emails is None:
        cc_emails = []

    pdf_file = Path(pdf_path).expanduser().resolve()
    if not pdf_file.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_file}")

    img_paths = pdf_to_images(str(pdf_file), dpi=dpi, max_pages=max_pages)

    pdf_file_escaped = Escape_AppleScript_Text(str(pdf_file))
    subject_escaped = Escape_AppleScript_Text(subject)
    body_text_escaped = Escape_AppleScript_Text(body_text)

    recipient_lines = "\n".join(
        f'make new recipient at end of to recipients of newMsg with properties {{email address:{{address:"{Escape_AppleScript_Text(email)}"}}}}'
        for email in to_emails
    )

    cc_lines = "\n".join(
    f'make new recipient at end of cc recipients of newMsg with properties {{email address:{{address:"{Escape_AppleScript_Text(email)}"}}}}'
    for email in cc_emails
)

    image_list_AppleScript = "{" + ", ".join(
        f'POSIX file "{Escape_AppleScript_Text(p)}" as alias' for p in img_paths
    ) + "}"

    text_block = '''
    delay 0.5
    tell application "System Events"
        tell process "Microsoft Outlook"
            set frontmost to true
            key code 36
            key code 36
            keystroke "b" using (command down)
            keystroke bodyText
            keystroke "b" using (command down)
        end tell
    end tell
    ''' if body_text else ""
    
    send_block = '''
    delay 2
    tell application "System Events"
        tell process "Microsoft Outlook"
            set frontmost to true
            keystroke return using {command down}
        end tell
    end tell
    ''' if send_automatically else ""

    hide_block = '''
    delay 0.5
    tell application "System Events"
        if exists process "Microsoft Outlook" then
            set visible of process "Microsoft Outlook" to false
        end if
    end tell
    '''
    
    AppleScript = f'''
                    set pdfFile to POSIX file "{pdf_file_escaped}" as alias
                    set msgSubject to "{subject_escaped}"
                    set bodyText to "{body_text_escaped}"
                    set imageFiles to {image_list_AppleScript}

                    tell application "Microsoft Outlook"
                        activate
                        set newMsg to make new outgoing message with properties {{subject:msgSubject}}
                        {recipient_lines}
                        {cc_lines}
                        open newMsg
                    end tell

                    delay 1.5

                    repeat with imgFile in imageFiles
                        set the clipboard to (read imgFile as picture)

                        tell application "System Events"
                            tell process "Microsoft Outlook"
                                set frontmost to true
                                keystroke "v" using command down

                            end tell
                        end tell

                        delay 0.8
                    end repeat

                    {text_block}

                    {send_block}

                    {hide_block}
                    '''

    result = subprocess.run(
        ["osascript", "-e", AppleScript],
        text=True,
        capture_output=True
    )

    if result.returncode != 0:
        raise RuntimeError(result.stderr)

    return {
        "returncode": result.returncode,
        "stdout": result.stdout,
        "stderr": result.stderr,
        "images_used": img_paths,
    }

