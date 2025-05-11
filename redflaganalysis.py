# redflaganalysis.py  OR  RedFlagAnalyzer.py

from transformers import pipeline

# Load the classifier from Hugging Face
classifier = pipeline("text-classification",
                      model="veejaydutt/bert-mdna-redflag",
                      tokenizer="veejaydutt/bert-mdna-redflag",
                      device=-1)  # Use device=0 if running on GPU

def get_red_flag_sentences(text):
    """
    Splits the input text into sentences and classifies each one.
    Returns only those classified as red flags (LABEL_1).
    """
    import re

    # Basic sentence split; improve later with nltk or spacy if needed
    sentences = re.split(r'(?<=[.?!])\s+', text.strip())

    red_flags = []
    for sentence in sentences:
        if not sentence.strip():
            continue
        result = classifier(sentence)[0]
        if result["label"] == "LABEL_1" and result["score"] > 0.6:
            red_flags.append(sentence.strip())
    return red_flags
