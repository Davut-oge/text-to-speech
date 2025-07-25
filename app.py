import os
import re
import argparse
import tempfile
import time
import shutil
import subprocess
import sys
import platform
from PyPDF2 import PdfReader
from gtts import gTTS
from pydub import AudioSegment
from tqdm import tqdm
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import logging

# Configure logging
logging.basicConfig(filename='pdf_to_audiobook.log', level=logging.ERROR,
                    format='%(asctime)s - %(levelname)s - %(message)s')


# Set ffmpeg path if not in system PATH
def setup_ffmpeg():
    """Ensure ffmpeg is available for pydub"""
    # First check if ffmpeg is in system PATH
    if shutil.which("ffmpeg"):
        return True

    # Try to find in common installation paths
    common_paths = []

    # Windows paths
    if platform.system() == "Windows":
        common_paths = [
            os.path.join(os.environ.get("PROGRAMFILES", "C:\\Program Files"), "ffmpeg", "bin", "ffmpeg.exe"),
            os.path.join(os.environ.get("PROGRAMFILES(X86)", "C:\\Program Files (x86)"), "ffmpeg", "bin", "ffmpeg.exe"),
            os.path.join(os.environ.get("SYSTEMDRIVE", "C:"), "ffmpeg", "bin", "ffmpeg.exe"),
            os.path.join(os.getcwd(), "ffmpeg.exe"),
        ]
    # Linux/Mac paths
    else:
        common_paths = [
            "/usr/bin/ffmpeg",
            "/usr/local/bin/ffmpeg",
            "/opt/bin/ffmpeg",
            os.path.join(os.getcwd(), "ffmpeg"),
        ]

    for path in common_paths:
        if os.path.exists(path):
            AudioSegment.converter = path
            return True

    return False


# Check ffmpeg availability at startup
FFMPEG_AVAILABLE = setup_ffmpeg()


def extract_text_from_pdf(pdf_path):
    """Extracts text from a PDF file"""
    text = ""
    try:
        with open(pdf_path, "rb") as file:
            reader = PdfReader(file)
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception as e:
        error_msg = f"PDF extraction failed: {str(e)}"
        logging.error(error_msg)
        raise RuntimeError(error_msg)
    return text


def clean_text(text):
    """Cleans and preprocesses extracted text"""
    try:
        # Remove excessive whitespace and newlines
        text = re.sub(r'\s+', ' ', text)
        # Remove non-printable characters
        text = re.sub(r'[^\x20-\x7E]', ' ', text)
        # Replace common PDF artifacts
        text = re.sub(r'\s*-\s*', '-', text)  # Fix hyphenated words
        # Remove page numbers and headers/footers
        text = re.sub(r'\bPage \d+\b', '', text)
        # Fix quotation marks
        text = re.sub(r'[\u201C\u201D]', '"', text)
        text = re.sub(r'[\u2018\u2019]', "'", text)
        return text.strip()
    except Exception as e:
        error_msg = f"Text cleaning failed: {str(e)}"
        logging.error(error_msg)
        raise RuntimeError(error_msg)


def split_text(text, max_chars=1000):
    """Splits text into chunks respecting sentence boundaries"""
    try:
        chunks = []
        while text:
            if len(text) <= max_chars:
                chunks.append(text)
                break

            # Find last sentence boundary within max_chars
            split_index = max_chars
            for boundary in ['.', '!', '?', ';', '\n', '。', '！', '？', '；']:
                index = text.rfind(boundary, 0, max_chars)
                if index > 0:
                    split_index = index + 1
                    break

            chunks.append(text[:split_index])
            text = text[split_index:].strip()

        return chunks
    except Exception as e:
        error_msg = f"Text splitting failed: {str(e)}"
        logging.error(error_msg)
        raise RuntimeError(error_msg)


def convert_text_to_speech(text, output_file, language='en'):
    """Converts text to speech using gTTS with progress handling"""
    try:
        if not FFMPEG_AVAILABLE:
            raise RuntimeError("ffmpeg is required but not found. Please install ffmpeg and add it to PATH.")

        # Split text into manageable chunks
        chunks = split_text(text)
        audio_segments = []
        temp_files = []  # Keep track of temp files for cleanup

        # Create progress window
        progress_window = tk.Toplevel()
        progress_window.title("Processing")
        progress_window.geometry("400x120")
        progress_window.resizable(False, False)
        progress_window.grab_set()

        progress_label = tk.Label(progress_window, text="Converting text to speech...")
        progress_label.pack(pady=10)

        progress = ttk.Progressbar(progress_window, orient="horizontal", length=300, mode="determinate")
        progress.pack(pady=5)

        status_label = tk.Label(progress_window, text="Preparing...")
        status_label.pack(pady=5)

        progress_window.update()

        try:
            for i, chunk in enumerate(chunks):
                # Update progress
                progress_val = int((i / len(chunks)) * 100)
                progress['value'] = progress_val
                status_label.config(text=f"Processing chunk {i + 1} of {len(chunks)}...")
                progress_window.update()

                # Convert text chunk to speech
                tts = gTTS(text=chunk, lang=language, slow=False)
                with tempfile.NamedTemporaryFile(delete=False, suffix=".mp3") as tmp:
                    temp_file_path = tmp.name
                    tts.save(temp_file_path)
                    temp_files.append(temp_file_path)

                # Check if file was created successfully
                if not os.path.exists(temp_file_path) or os.path.getsize(temp_file_path) == 0:
                    raise RuntimeError(f"Failed to create temporary audio file: {temp_file_path}")

                audio_segments.append(AudioSegment.from_mp3(temp_file_path))

            # Update progress for combining audio
            status_label.config(text="Combining audio segments...")
            progress['value'] = 100
            progress_window.update()

            # Combine audio segments
            combined_audio = audio_segments[0]
            for segment in audio_segments[1:]:
                combined_audio += segment

            # Export final audio
            combined_audio.export(output_file, format="mp3")

            # Check if output file was created
            if not os.path.exists(output_file) or os.path.getsize(output_file) == 0:
                raise RuntimeError(f"Failed to create output file: {output_file}")

            return True
        except Exception as e:
            error_msg = f"Text-to-speech conversion failed: {str(e)}"
            logging.error(error_msg)
            raise
        finally:
            # Clean up temporary files
            for temp_file in temp_files:
                try:
                    if os.path.exists(temp_file):
                        os.remove(temp_file)
                except Exception as e:
                    logging.warning(f"Failed to delete temp file {temp_file}: {str(e)}")

            progress_window.destroy()
    except Exception as e:
        error_msg = f"Text-to-speech conversion failed: {str(e)}"
        logging.error(error_msg)
        raise RuntimeError(error_msg)


def convert_pdf_to_speech(pdf_path, output_file, language='en'):
    """Full pipeline: PDF to text to speech"""
    try:
        # Check input file
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")

        # Extract and clean text
        raw_text = extract_text_from_pdf(pdf_path)
        if not raw_text:
            return False, "No text extracted from PDF. The file may be scanned or image-based."

        cleaned_text = clean_text(raw_text)

        # Convert to speech
        convert_text_to_speech(cleaned_text, output_file, language)
        return True, "Conversion successful!"
    except Exception as e:
        error_msg = f"PDF to speech conversion failed: {str(e)}"
        logging.error(error_msg)
        return False, error_msg


def setup_gui():
    """Create and configure the GUI"""
    root = tk.Tk()
    root.title("PDF to Audiobook Converter")
    root.geometry("700x500")
    root.resizable(True, True)

    # Set application icon
    try:
        icon_path = 'icon.ico'
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
    except Exception as e:
        logging.warning(f"Could not set application icon: {str(e)}")

    # Configure styles
    style = ttk.Style()
    style.configure('TButton', padding=6, relief="flat", background="#4CAF50")
    style.configure('TFrame', background='#f0f0f0')

    # Create main frame
    main_frame = ttk.Frame(root, padding="20")
    main_frame.pack(fill=tk.BOTH, expand=True)

    # PDF selection
    pdf_frame = ttk.Frame(main_frame)
    pdf_frame.pack(fill=tk.X, pady=(0, 10))

    ttk.Label(pdf_frame, text="PDF File:").pack(side=tk.LEFT)
    pdf_path_var = tk.StringVar()
    pdf_entry = ttk.Entry(pdf_frame, textvariable=pdf_path_var, width=50, state='readonly')
    pdf_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

    def browse_pdf():
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            try:
                # Verify file exists
                if not os.path.exists(file_path):
                    messagebox.showerror("Error", f"File not found: {file_path}")
                    return

                raw_text = extract_text_from_pdf(file_path)
                if not raw_text:
                    messagebox.showwarning("Warning", "No text extracted from PDF. The file may be scanned.")
                    return

                cleaned_text = clean_text(raw_text)
                text_input.delete("1.0", tk.END)
                text_input.insert(tk.END, cleaned_text)
                pdf_path_var.set(file_path)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load PDF: {str(e)}")

    ttk.Button(pdf_frame, text="Browse", command=browse_pdf).pack(side=tk.LEFT)

    # Text input
    ttk.Label(main_frame, text="Text Content:").pack(anchor=tk.W)
    text_frame = ttk.Frame(main_frame)
    text_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

    text_scroll = ttk.Scrollbar(text_frame)
    text_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    text_input = tk.Text(text_frame, wrap=tk.WORD, yscrollcommand=text_scroll.set)
    text_input.pack(fill=tk.BOTH, expand=True)
    text_scroll.config(command=text_input.yview)

    # Options
    options_frame = ttk.Frame(main_frame)
    options_frame.pack(fill=tk.X, pady=(0, 10))

    ttk.Label(options_frame, text="Language:").pack(side=tk.LEFT)
    lang_var = tk.StringVar(value='en')
    lang_menu = ttk.Combobox(options_frame, textvariable=lang_var, width=10, state='readonly')
    lang_menu['values'] = (
        'en', 'es', 'fr', 'de', 'it', 'pt', 'ru', 'tr',
        'ar', 'zh-CN', 'ja', 'hi', 'ko', 'nl', 'sv', 'pl'
    )
    lang_menu.pack(side=tk.LEFT, padx=5)

    # Voice speed option
    ttk.Label(options_frame, text="Speed:").pack(side=tk.LEFT, padx=(20, 0))
    speed_var = tk.DoubleVar(value=1.0)
    speed_scale = ttk.Scale(options_frame, from_=0.5, to=2.0, variable=speed_var,
                            orient=tk.HORIZONTAL, length=80)
    speed_scale.pack(side=tk.LEFT, padx=5)
    speed_value_label = ttk.Label(options_frame, text="1.0x")
    speed_value_label.pack(side=tk.LEFT, padx=5)

    def update_speed_label(*args):
        speed_value_label.config(text=f"{speed_var.get():.1f}x")

    speed_var.trace_add("write", update_speed_label)

    # Play after conversion option
    play_var = tk.BooleanVar(value=True)
    ttk.Checkbutton(options_frame, text="Play after conversion", variable=play_var).pack(side=tk.LEFT, padx=20)

    # Action buttons
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill=tk.X)

    def generate_speech():
        """Handle speech generation from GUI"""
        # Get the text from the text widget - this includes any user edits
        text = text_input.get("1.0", tk.END).strip()
        pdf_path = pdf_path_var.get()

        if not text and not pdf_path:
            messagebox.showwarning("Warning", "Please enter text or load a PDF file!")
            return

        # Get output file path
        output_path = filedialog.asksaveasfilename(
            defaultextension=".mp3",
            filetypes=[("MP3 Files", "*.mp3")]
        )
        if not output_path:
            return

        # Create output directory if needed
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create output directory: {str(e)}")
                return

        # Get language selection
        language = lang_var.get()

        try:
            # Always use the text from the text widget (which may include user edits)
            convert_text_to_speech(text, output_path, language)
            success, message = True, "Conversion successful!"

            # Apply speed adjustment
            speed_factor = speed_var.get()
            if speed_factor != 1.0:
                try:
                    audio = AudioSegment.from_mp3(output_path)
                    new_sample_rate = int(audio.frame_rate * speed_factor)
                    adjusted_audio = audio._spawn(audio.raw_data, overrides={
                        'frame_rate': new_sample_rate
                    })
                    adjusted_audio = adjusted_audio.set_frame_rate(audio.frame_rate)
                    adjusted_audio.export(output_path, format="mp3")
                except Exception as e:
                    messagebox.showwarning("Speed Adjustment", f"Speed adjustment failed: {str(e)}")

            messagebox.showinfo("Success", f"{message}\nFile saved to: {output_path}")
            if play_var.get():
                if sys.platform == 'win32':
                    os.startfile(output_path)  # Windows
                elif sys.platform == 'darwin':
                    subprocess.call(('open', output_path))  # macOS
                else:
                    subprocess.call(('xdg-open', output_path))  # Linux
        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")

    ttk.Button(button_frame, text="Convert to Audiobook", command=generate_speech,
               style='TButton').pack(pady=10)

    # Status bar
    status_bar = ttk.Label(root, text="Ready", relief=tk.SUNKEN, anchor=tk.W)
    status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    # FFmpeg status
    if not FFMPEG_AVAILABLE:
        status_bar.config(text="Warning: ffmpeg not found - audio processing may not work",
                          background="#ffcc00", foreground="black")

    # Add FFmpeg download link
    if not FFMPEG_AVAILABLE and platform.system() == "Windows":
        def open_ffmpeg_download():
            import webbrowser
            webbrowser.open("https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip")

        ffmpeg_frame = ttk.Frame(main_frame)
        ffmpeg_frame.pack(fill=tk.X, pady=5)
        ttk.Label(ffmpeg_frame, text="FFmpeg is required for audio processing").pack(side=tk.LEFT)
        ttk.Button(ffmpeg_frame, text="Download FFmpeg", command=open_ffmpeg_download).pack(side=tk.RIGHT)

    return root


def main_gui():
    """Run the GUI application"""
    try:
        root = setup_gui()
        root.mainloop()
    except Exception as e:
        logging.exception("Critical error in GUI")
        messagebox.showerror("Fatal Error", f"A critical error occurred: {str(e)}\nSee log file for details.")


def main_cli():
    """Command-line interface functionality"""
    if not FFMPEG_AVAILABLE:
        print("Warning: ffmpeg not found - audio processing may not work")

    parser = argparse.ArgumentParser(description="Convert PDF to audiobook")
    parser.add_argument("pdf_file", help="Path to input PDF file")
    parser.add_argument("output_file", help="Output MP3 file path")
    parser.add_argument("-l", "--lang", default="en", help="Language code (default: en)")
    parser.add_argument("-s", "--speed", type=float, default=1.0, help="Playback speed (0.5 to 2.0, default: 1.0)")
    args = parser.parse_args()

    # Check input file
    if not os.path.exists(args.pdf_file):
        print(f"Error: PDF file not found at {args.pdf_file}")
        return

    # Create output directory if needed
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
        except Exception as e:
            print(f"Error: Failed to create output directory: {str(e)}")
            return

    print("Extracting text from PDF...")
    success, message = convert_pdf_to_speech(args.pdf_file, args.output_file, args.lang)

    if success:
        # Apply speed adjustment if needed
        if args.speed != 1.0:
            try:
                print("Adjusting playback speed...")
                audio = AudioSegment.from_mp3(args.output_file)
                new_sample_rate = int(audio.frame_rate * args.speed)
                adjusted_audio = audio._spawn(audio.raw_data, overrides={
                    'frame_rate': new_sample_rate
                })
                adjusted_audio = adjusted_audio.set_frame_rate(audio.frame_rate)
                adjusted_audio.export(args.output_file, format="mp3")
            except Exception as e:
                print(f"Warning: Speed adjustment failed: {str(e)}")

        print(f"Success: {message}")
        print(f"Audiobook saved to: {os.path.abspath(args.output_file)}")
    else:
        print(f"Error: {message}")


if __name__ == "__main__":
    # Run in CLI mode if arguments are passed, otherwise run GUI
    if len(sys.argv) > 1:
        main_cli()
    else:
        main_gui()