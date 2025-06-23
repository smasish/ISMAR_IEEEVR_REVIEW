import os
from spire.pdf import *
from spire.pdf.common import *

def highlight_terms_in_pdf(input_path, output_path, terms_list, highlight_color=None):
    """
    Highlight specified terms in a PDF file and save the result.

    Args:
        input_path (str): Path to the input PDF file
        output_path (str): Path to save the highlighted PDF
        terms_list (list): List of terms to highlight
        highlight_color (Color, optional): Color for highlighting. Default is yellow.
    """
    try:
        # Load the PDF document
        pdf = PdfDocument()
        pdf.LoadFromFile(input_path)

        print(f"Processing: {os.path.basename(input_path)} - {pdf.Pages.Count} pages")

        # Process each page
        for i in range(pdf.Pages.Count):
            page = pdf.Pages.get_Item(i)

            # Search for each term using the original approach
            for term in terms_list:
                # Create variations of the term for better matching
                term_variations = [
                    f" {term} ",  # Term with spaces on both sides
                    f"{term} ",   # Term with space after
                    f" {term}",   # Term with space before
                    f"{term},",   # Term followed by comma
                    f"({term})",  # Term in parentheses
                    f"{term}.",   # Term followed by period
                    f"{term}:",   # Term followed by colon
                    f"{term};"    # Term followed by semicolon
                ]

                # Method 1: Using PdfTextFinder (if available)
                try:
                    finder = PdfTextFinder(page)
                    finder.Options.Parameter = TextFindParameter.IgnoreCase

                    for variation in term_variations:
                        text_fragments = finder.Find(variation)
                        for fragment in text_fragments:
                            if highlight_color:
                                fragment.ApplyHighLight(highlight_color)
                            else:
                                fragment.HighLight()

                except (NameError, AttributeError):
                    # Method 2: Using FindText method (fallback)
                    for variation in term_variations:
                        try:
                            # Try different parameter approaches
                            result = None

                            # Try with TextFindParameter if available
                            try:
                                result = page.FindText(variation, TextFindParameter.IgnoreCase).Finds
                            except NameError:
                                # Try with string parameter
                                try:
                                    result = page.FindText(variation, "IgnoreCase").Finds
                                except:
                                    # Try without parameter
                                    result = page.FindText(variation).Finds

                            # Highlight found text
                            if result:
                                for text in result:
                                    if highlight_color:
                                        text.ApplyHighLight(highlight_color)
                                    else:
                                        text.ApplyHighLight(Color.get_Yellow())

                        except Exception as e:
                            continue  # Skip this variation if it fails

        # Save the highlighted PDF
        pdf.SaveToFile(output_path)
        pdf.Close()
        print(f"Saved highlighted PDF to: {output_path}")
        return True

    except Exception as e:
        print(f"Error processing {input_path}: {str(e)}")
        return False

def process_directory(input_dir, output_dir, terms_list, highlight_color=None):
    """
    Process all PDF files in a directory.

    Args:
        input_dir (str): Directory containing PDF files
        output_dir (str): Directory to save highlighted PDFs
        terms_list (list): List of terms to highlight
        highlight_color (Color, optional): Color for highlighting
    """
    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Get all PDF files in the input directory
    pdf_files = [f for f in os.listdir(input_dir) if f.lower().endswith('.pdf')]

    if not pdf_files:
        print(f"No PDF files found in {input_dir}")
        return

    print(f"Found {len(pdf_files)} PDF files to process")

    # Process each PDF file
    successful = 0
    for pdf_file in pdf_files:
        input_path = os.path.join(input_dir, pdf_file)
        output_path = os.path.join(output_dir, f"highlighted_{pdf_file}")

        if highlight_terms_in_pdf(input_path, output_path, terms_list, highlight_color):
            successful += 1

    print(f"Successfully processed {successful} out of {len(pdf_files)} PDF files")

# Alternative simplified version based on your original working code
def highlight_terms_simple(input_path, output_path, terms_list):
    """
    Simplified version based on your original working code structure.
    """
    try:
        pdf = PdfDocument()
        pdf.LoadFromFile(input_path)

        print(f"Processing: {os.path.basename(input_path)} - {pdf.Pages.Count} pages")

        # Loop through the pages in the PDF document
        for i in range(pdf.Pages.Count):
            for keyword in terms_list:
                # Get a page
                page = pdf.Pages.get_Item(i)

                # Create a PdfTextFinder object for the current page
                finder = PdfTextFinder(page)

                # Set the search parameter to find exact matches
                finder.Options.Parameter = TextFindParameter.IgnoreCase

                # Find instances of a sentence on the page
                modified_keywords = [f' {keyword} ', f'{keyword} ', f'{keyword},', f'({keyword})', ]

                for mkeyword in modified_keywords:
                    textFragments = finder.Find(mkeyword)

                    # Iterate through the instances
                    for textFragment in textFragments:
                        # Highlight each instance
                        textFragment.HighLight()

        # Save the document
        pdf.SaveToFile(output_path)
        pdf.Close()
        print(f"Saved highlighted PDF to: {output_path}")
        return True

    except Exception as e:
        print(f"Error processing {input_path}: {str(e)}")
        return False

# Example usage
if __name__ == "__main__":
    # Define directories
    input_directory = r"C:\Users\Shaon\Downloads\Documents\Papers\Original"
    output_directory = r"C:\Users\Shaon\Downloads\Documents\Papers\Highlighted"

    # Define terms to highlight
    ml_terms = [
        "machine learning", "ml", "deep learning", "dl",
        "artificial", "classification",
        "prediction", "classifier", "detection", "semantic", "segmentation", "regression",
        "logistic", "random forest", "rf", "support vector", "svm", "generative",
        "adversarial", "gan", "perceptron", "neuron", "neural", "network", "ann",
        "convolution", "convolutional", "cnn", "reinforcement", "rl", "radial", "rbf",
        "attention", "transformer", "segmentation", "k-nearest", "knn", "boosting",
        "xgboost", "tree", "graph", "least", "least square", "short-term", "lstm",
        "gru", "hmm", "rnn", "autoregressive", "ridge", "lasso"
    ]

    # Process all PDFs in the directory using the simplified version
    pdf_files = [f for f in os.listdir(input_directory) if f.lower().endswith('.pdf')]

    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    for pdf_file in pdf_files:
        input_path = os.path.join(input_directory, pdf_file)
        output_path = os.path.join(output_directory, f"highlighted_{pdf_file}")
        highlight_terms_simple(input_path, output_path, ml_terms)
