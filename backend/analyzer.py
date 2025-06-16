import re
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.metrics.pairwise import cosine_similarity

def analyze_cvs(filepaths, job_requirements):
    # Extract text from files
    from backend.file_handler import extract_text_from_files
    texts = extract_text_from_files(filepaths)

    # Include job requirements in the analysis
    all_texts = [job_requirements] + texts

    # Vectorize the texts
    vectorizer = CountVectorizer().fit_transform(all_texts)
    vectors = vectorizer.toarray()

    # Compute similarity scores
    similarity_matrix = cosine_similarity(vectors)
    scores = similarity_matrix[0][1:]  # Compare job requirements with each CV

    # Rank candidates
    candidates = [filepath.split("/")[-1] for filepath in filepaths]
    ranked_results = sorted(zip(candidates, scores), key=lambda x: x[1], reverse=True)

    return ranked_results
