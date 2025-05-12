# redflaganalysis.py  OR  RedFlagAnalyzer.py

from transformers import pipeline

# Load the classifier from Hugging Face
classifier = pipeline("text-classification",
                      model="veejaydutt/bert-mdna-redflag",
                      tokenizer="veejaydutt/bert-mdna-redflag",
                      device=-1)  # Use device=0 if running on GPU

def get_red_flag_sentences(text):
    import re
    if text.lower().startswith("highlight red flag statements in this passage:"):
        text = text[len("highlight red flag statements in this passage:"):].strip()

    sentences = re.split(r'(?<=[.?!])\s+', text.strip())
    red_flags = []

    for sentence in sentences:
        if not sentence.strip():
            continue
        result = classifier(sentence)[0]
        if result["label"] == "LABEL_1" and result["score"] > 0.75:
            red_flags.append((sentence.strip(), round(result["score"], 2)))

    return red_flags


