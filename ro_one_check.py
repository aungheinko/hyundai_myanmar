from sklearn.feature_extraction.text import CountVectorizer
from sklearn.linear_model import LinearRegression

# Example sentences and their associated scores
texts = ["I love this!", "This is amazing!", "I hate it.", "This is terrible.", "It's okay."]
scores = [9, 10, 2, 1, 5]  # Higher score means more positive

# Convert text to numerical data
vectorizer = CountVectorizer()
X = vectorizer.fit_transform(texts)

# Create a linear regression model
model = LinearRegression()
model.fit(X, scores)

# Predict the score of a new sentence
new_text = ["This is fantastic!"]
new_X = vectorizer.transform(new_text)
predicted_score = model.predict(new_X)

print(f"The predicted score for '{new_text[0]}' is: {predicted_score[0]:.2f}")
