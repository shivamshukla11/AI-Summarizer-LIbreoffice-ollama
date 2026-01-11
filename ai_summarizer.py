import subprocess
import uno
import time
import json

def load_config():
    """Loads the configuration from the config.json file."""
    try:
        with open("config.json", "r") as f:
            config = json.load(f)
            return config.get("OLLAMA_PATH"), config.get("MODEL")
    except FileNotFoundError:
        return None, None

def get_selected_text():
    """Gets the selected text from the current LibreOffice document."""
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.getCurrentComponent()

    if not hasattr(doc, "Text"):
        return None, None, None

    controller = doc.getCurrentController()
    selection = controller.getSelection()
    if selection.getCount() == 0:
        return None, None, None

    text_range = selection.getByIndex(0)
    input_text = text_range.getString().strip()
    return doc, text_range, input_text

def show_progress_indicator(doc, text_range):
    """Inserts a progress indicator in the document."""
    cursor = doc.Text.createTextCursorByRange(text_range.getEnd())
    doc.Text.insertString(cursor, "\n\n[AI is summarizing… please wait]\n", False)
    time.sleep(0.2)  # Let LibreOffice render before blocking
    return cursor

def invoke_ollama(input_text, ollama_path, model):
    """Invokes the Ollama AI model to summarize the text."""
    prompt = f"Summarize the following text clearly:\n\n{input_text}"
    try:
        process = subprocess.run(
            [ollama_path, "run", model],
            input=prompt,
            capture_output=True,
            text=True,
            encoding="utf-8",
            check=True  # Raise an exception for non-zero exit codes
        )
        return process.stdout.strip()
    except FileNotFoundError:
        return f"[ERROR] Ollama executable not found at: {ollama_path}"
    except subprocess.CalledProcessError as e:
        return f"[ERROR] Ollama returned an error:\n{e.stderr.strip()}"
    except Exception as e:
        return f"[ERROR] An unexpected error occurred: {e}"

def display_summary(cursor, summary):
    """Displays the summary in the document."""
    cursor.setString("\n\n--- Summary (AI) ---\n" + summary + "\n")

def summarize_text():
    """Main function to summarize the selected text."""
    doc, text_range, input_text = get_selected_text()

    if not doc:
        return  # No document, can't proceed

    ollama_path, model = load_config()
    if not ollama_path or not model:
        # If config is missing, we can't show a progress indicator,
        # so we must display the error differently.
        error_cursor = doc.Text.createTextCursor()
        error_cursor.gotoEnd(False)
        doc.Text.insertString(error_cursor, "\n\n[ERROR] Configuration is missing.", False)
        return

    if not input_text:
        return  # No text selected, do nothing.

    cursor = show_progress_indicator(doc, text_range)
    summary = invoke_ollama(input_text, ollama_path, model)

    # Display either the summary or an error using the same cursor
    display_summary(cursor, summary)
