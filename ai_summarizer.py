import subprocess
import uno
import time

OLLAMA_PATH = r"C:\Users\shiva\AppData\Local\Programs\Ollama\ollama.exe"
# MODEL = "llama2"
# MODEL = "mistral"
MODEL = "phi"

def summarize_text():
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext(
        "com.sun.star.frame.Desktop", ctx
    )

    doc = desktop.getCurrentComponent()
    if not hasattr(doc, "Text"):
        return

    controller = doc.getCurrentController()
    selection = controller.getSelection()
    if selection.getCount() == 0:
        return

    text_range = selection.getByIndex(0)
    input_text = text_range.getString().strip()
    if not input_text:
        return

    # Insert progress marker (safe, no invalidate)
    cursor = doc.Text.createTextCursorByRange(text_range.getEnd())
    doc.Text.insertString(
        cursor,
        "\n\n[AI is summarizing… please wait]\n",
        False
    )

    # Let LibreOffice render before blocking
    time.sleep(0.2)

    prompt = f"Summarize the following text clearly:\n\n{input_text}"

    process = subprocess.run(
        [OLLAMA_PATH, "run", MODEL],
        input=prompt,
        capture_output=True,
        text=True,
        encoding="utf-8"
    )

    output = process.stdout.strip()
    if not output:
        output = process.stderr.strip() or "[ERROR] Ollama returned no output"

    # Replace progress text with result
    cursor.setString(
        "\n\n--- Summary (AI) ---\n" + output + "\n"
    )
