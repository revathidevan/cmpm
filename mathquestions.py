from flask import Flask, request, jsonify
import random

app = Flask(__name__)

# Collections of questions
square_area_questions = [
    {"question": "What is the area of a square with side length 4?", "answer": 16},
    {"question": "What is the area of a square with side length 7?", "answer": 49},
    {"question": "What is the area of a square with side length 3?", "answer": 9},
    {"question": "What is the area of a square with side length 5?", "answer": 25},
    {"question": "What is the area of a square with side length 6?", "answer": 36},
    {"question": "What is the area of a square with side length 8?", "answer": 64},
    {"question": "What is the area of a square with side length 2?", "answer": 4},
    {"question": "What is the area of a square with side length 9?", "answer": 81},
    {"question": "What is the area of a square with side length 10?", "answer": 100},
    {"question": "What is the area of a square with side length 1?", "answer": 1},
]

square_perimeter_questions = [
    {"question": "What is the perimeter of a square with side length 4?", "answer": 16},
    {"question": "What is the perimeter of a square with side length 7?", "answer": 28},
    {"question": "What is the perimeter of a square with side length 3?", "answer": 12},
    {"question": "What is the perimeter of a square with side length 5?", "answer": 20},
    {"question": "What is the perimeter of a square with side length 6?", "answer": 24},
    {"question": "What is the perimeter of a square with side length 8?", "answer": 32},
    {"question": "What is the perimeter of a square with side length 2?", "answer": 8},
    {"question": "What is the perimeter of a square with side length 9?", "answer": 36},
    {"question": "What is the perimeter of a square with side length 10?", "answer": 40},
    {"question": "What is the perimeter of a square with side length 1?", "answer": 4},
]

rectangle_area_questions = [
    {"question": "What is the area of a rectangle with length 4 and width 3?", "answer": 12},
    {"question": "What is the area of a rectangle with length 7 and width 5?", "answer": 35},
    {"question": "What is the area of a rectangle with length 6 and width 2?", "answer": 12},
    {"question": "What is the area of a rectangle with length 8 and width 4?", "answer": 32},
    {"question": "What is the area of a rectangle with length 5 and width 5?", "answer": 25},
    {"question": "What is the area of a rectangle with length 9 and width 3?", "answer": 27},
    {"question": "What is the area of a rectangle with length 10 and width 2?", "answer": 20},
    {"question": "What is the area of a rectangle with length 3 and width 3?", "answer": 9},
    {"question": "What is the area of a rectangle with length 6 and width 4?", "answer": 24},
    {"question": "What is the area of a rectangle with length 7 and width 3?", "answer": 21},
]

rectangle_perimeter_questions = [
    {"question": "What is the perimeter of a rectangle with length 4 and width 3?", "answer": 14},
    {"question": "What is the perimeter of a rectangle with length 7 and width 5?", "answer": 24},
    {"question": "What is the perimeter of a rectangle with length 6 and width 2?", "answer": 16},
    {"question": "What is the perimeter of a rectangle with length 8 and width 4?", "answer": 24},
    {"question": "What is the perimeter of a rectangle with length 5 and width 5?", "answer": 20},
    {"question": "What is the perimeter of a rectangle with length 9 and width 3?", "answer": 24},
    {"question": "What is the perimeter of a rectangle with length 10 and width 2?", "answer": 24},
    {"question": "What is the perimeter of a rectangle with length 3 and width 3?", "answer": 12},
    {"question": "What is the perimeter of a rectangle with length 6 and width 4?", "answer": 20},
    {"question": "What is the perimeter of a rectangle with length 7 and width 3?", "answer": 20},
]

triangle_area_questions = [
    {"question": "What is the area of a triangle with base 4 and height 3?", "answer": 6},
    {"question": "What is the area of a triangle with base 6 and height 5?", "answer": 15},
    {"question": "What is the area of a triangle with base 8 and height 4?", "answer": 16},
    {"question": "What is the area of a triangle with base 5 and height 7?", "answer": 17.5},
    {"question": "What is the area of a triangle with base 10 and height 3?", "answer": 15},
    {"question": "What is the area of a triangle with base 7 and height 6?", "answer": 21},
    {"question": "What is the area of a triangle with base 9 and height 4?", "answer": 18},
    {"question": "What is the area of a triangle with base 3 and height 8?", "answer": 12},
    {"question": "What is the area of a triangle with base 6 and height 6?", "answer": 18},
    {"question": "What is the area of a triangle with base 5 and height 5?", "answer": 12.5},
]

triangle_perimeter_questions = [
    {"question": "What is the perimeter of a triangle with sides 3, 4, and 5?", "answer": 12},
    {"question": "What is the perimeter of a triangle with sides 5, 5, and 7?", "answer": 17},
    {"question": "What is the perimeter of a triangle with sides 6, 8, and 10?", "answer": 24},
    {"question": "What is the perimeter of a triangle with sides 4, 4, and 4?", "answer": 12},
    {"question": "What is the perimeter of a triangle with sides 7, 8, and 9?", "answer": 24},
    {"question": "What is the perimeter of a triangle with sides 5, 12, and 13?", "answer": 30},
    {"question": "What is the perimeter of a triangle with sides 3, 3, and 3?", "answer": 9},
    {"question": "What is the perimeter of a triangle with sides 8, 8, and 8?", "answer": 24},
    {"question": "What is the perimeter of a triangle with sides 6, 6, and 6?", "answer": 18},
    {"question": "What is the perimeter of a triangle with sides 10, 10, and 10?", "answer": 30},
]

@app.route('/questions', methods=['GET'])
def get_questions():
    shape = request.args.get('shape', '').lower()
    method = request.args.get('method', '').lower()
    num_questions = int(request.args.get('num', 1))

    if shape not in ['square', 'rectangle', 'triangle'] or method not in ['area', 'perimeter']:
        return jsonify({"error": "Invalid shape or method"}), 400

    question_set = globals()[f"{shape}_{method}_questions"]
    selected_questions = random.sample(question_set, min(num_questions, len(question_set)))

    return jsonify(selected_questions)

@app.route('/submit', methods=['POST'])
def submit_answers():
    data = request.json
    if not data or 'answers' not in data:
        return jsonify({"error": "No answers provided"}), 400

    answers = data['answers']
    score = 0
    total = len(answers)

    for answer in answers:
        if 'question' in answer and 'user_answer' in answer and 'correct_answer' in answer:
            if float(answer['user_answer']) == float(answer['correct_answer']):
                score += 1

    return jsonify({
        "score": score,
        "total": total,
        "percentage": (score / total) * 100 if total > 0 else 0
    })

@app.route('/', methods=['GET'])
def home():
    return "Welcome to the Math Questions API. Use /questions to get questions and /submit to submit answers."

if __name__ == '__main__':
    app.run(debug=True)
