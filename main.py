import gradio as gr
import pypandoc
import tempfile
import os
from pathlib import Path


def convert_md_to_word(markdown_text):
    """Convert markdown text to Word document using pypandoc"""
    try:
        # Create persistent temp file
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as temp_file:
            output_path = temp_file.name

        # Create temp markdown input
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".md", encoding="utf-8", delete=False
        ) as md_file:
            md_file.write(markdown_text)
            md_path = md_file.name

        # Convert using pandoc
        pypandoc.convert_file(
            md_path, "docx", outputfile=output_path, format="markdown"
        )

        # Cleanup input file
        os.unlink(md_path)

        # Return path with Gradio file validation
        return gr.File(value=output_path, visible=True)

    except Exception as e:
        # Cleanup any remaining temp files
        if "md_path" in locals() and os.path.exists(md_path):
            os.unlink(md_path)
        if "output_path" in locals() and os.path.exists(output_path):
            os.unlink(output_path)
        raise gr.Error(f"Conversion failed: {str(e)}")


# Create Gradio interface with explicit allowed paths
interface = gr.Interface(
    fn=convert_md_to_word,
    inputs=gr.Textbox(
        label="Markdown Input", placeholder="Enter your markdown text here...", lines=20
    ),
    outputs=gr.File(label="Download Word Document"),
    title="Markdown to Word Converter",
    description="Convert markdown text to Microsoft Word document (DOCX)",
    allow_flagging="never",
)

if __name__ == "__main__":
    interface.launch(
        allowed_paths=[tempfile.gettempdir()]  # Grant access to temp directory
    )
