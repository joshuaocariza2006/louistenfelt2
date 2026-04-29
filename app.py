from pathlib import Path
import re
from flask import Flask, render_template, request, flash, redirect, url_for, session, send_from_directory
from openpyxl import Workbook, load_workbook

app = Flask(__name__)
app.secret_key = 'change-this-secret-key'

DATA_FILE = Path(__file__).parent / 'users.xlsx'
EMAIL_RE = re.compile(r'^[^@\s]+@[^@\s]+\.[^@\s]+$')
EXPECTED_HEADERS = ['username', 'email', 'password', 'display_name', 'role', 'tagline', 'likes', 'dislikes', 'birthday', 'place', 'fav_games']


def ensure_data_file():
    if not DATA_FILE.exists():
        wb = Workbook()
        ws = wb.active
        ws.title = 'Users'
        ws.append(EXPECTED_HEADERS)
        wb.save(DATA_FILE)
        return

    workbook = load_workbook(DATA_FILE)
    sheet = workbook.active
    existing_headers = [str(cell.value).strip() if cell.value is not None else '' for cell in next(sheet.iter_rows(min_row=1, max_row=1, values_only=False))]
    changed = False
    for column_index, header_name in enumerate(EXPECTED_HEADERS, start=1):
        if column_index > len(existing_headers) or existing_headers[column_index - 1] != header_name:
            sheet.cell(row=1, column=column_index, value=header_name)
            changed = True
    if changed:
        workbook.save(DATA_FILE)


def load_users():
    ensure_data_file()
    workbook = load_workbook(DATA_FILE)
    sheet = workbook.active
    headers = [str(cell.value).strip() if cell.value is not None else '' for cell in next(sheet.iter_rows(min_row=1, max_row=1, values_only=False))]
    users = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not row or not any(cell is not None and str(cell).strip() != '' for cell in row):
            continue
        row_data = {
            headers[i] if i < len(headers) else f'col_{i}': str(row[i]).strip() if row[i] is not None else ''
            for i in range(len(row))
        }
        users.append({
            'username': row_data.get('username', ''),
            'email': row_data.get('email', ''),
            'password': row_data.get('password', ''),
            'display_name': row_data.get('display_name', ''),
            'role': row_data.get('role', ''),
            'tagline': row_data.get('tagline', ''),
            'likes': row_data.get('likes', ''),
            'dislikes': row_data.get('dislikes', ''),
            'birthday': row_data.get('birthday', ''),
            'place': row_data.get('place', ''),
            'fav_games': row_data.get('fav_games', ''),
        })
    return users


def save_user(username: str, email: str, password: str, display_name: str = '', role: str = '', tagline: str = '', likes: str = '', dislikes: str = '', birthday: str = '', place: str = '', fav_games: str = ''):
    ensure_data_file()
    workbook = load_workbook(DATA_FILE)
    sheet = workbook.active
    sheet.append([username, email, password, display_name, role, tagline, likes, dislikes, birthday, place, fav_games])
    workbook.save(DATA_FILE)


def update_user(current_username: str, new_data: dict):
    ensure_data_file()
    workbook = load_workbook(DATA_FILE)
    sheet = workbook.active
    headers = [str(cell.value).strip() if cell.value is not None else '' for cell in next(sheet.iter_rows(min_row=1, max_row=1, values_only=False))]
    updated = False
    for row in sheet.iter_rows(min_row=2):
        cell_value = row[0].value
        if cell_value is None:
            continue
        if str(cell_value).strip().lower() != current_username.strip().lower():
            continue
        for key, value in new_data.items():
            if key in headers:
                col_index = headers.index(key)
                row[col_index].value = value
        updated = True
        break
    if updated:
        workbook.save(DATA_FILE)
    return updated


def find_user(username_or_email: str):
    value = username_or_email.strip().lower()
    users = load_users()
    for user in users:
        if user['username'].lower() == value or user['email'].lower() == value:
            return user
    return None


def get_user_by_username(username: str):
    if not username:
        return None
    users = load_users()
    for user in users:
        if user['username'].lower() == username.strip().lower():
            return user
    return None


def get_current_user():
    return get_user_by_username(session.get('username', ''))


@app.route('/', methods=['GET', 'POST'])
def auth():
    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'register':
            username = request.form.get('username', '').strip()
            email = request.form.get('email', '').strip()
            password = request.form.get('password', '')
            confirm_password = request.form.get('confirm_password', '')

            if not username or not email or not password or not confirm_password:
                flash('Please fill in every registration field.', 'error')
            elif not EMAIL_RE.match(email):
                flash('Please enter a valid email address.', 'error')
            elif password != confirm_password:
                flash('Passwords do not match.', 'error')
            elif len(password) < 6:
                flash('Password must be at least 6 characters long.', 'error')
            elif find_user(username) is not None:
                flash('A user with that username or email already exists.', 'error')
            elif find_user(email) is not None:
                flash('A user with that username or email already exists.', 'error')
            else:
                save_user(username, email, password)
                flash('Registration successful! You can now log in.', 'success')
                return redirect(url_for('auth') + '#login')

        elif action == 'login':
            username_or_email = request.form.get('username_or_email', '').strip()
            password = request.form.get('password', '')

            if not username_or_email or not password:
                flash('Please enter username/email and password to log in.', 'error')
            else:
                user = find_user(username_or_email)
                if user is None or user['password'] != password:
                    flash('Invalid login credentials.', 'error')
                else:
                    session['username'] = user['username']
                    flash(f'Login successful. Welcome, {user["username"]}!', 'success')
                    return redirect(url_for('homepage'))

        else:
            flash('Invalid form submission.', 'error')

    return render_template('index.html')


@app.route('/homepage')
def homepage():
    user = get_current_user()
    if not user:
        flash('Please log in to access the homepage.', 'error')
        session.pop('username', None)
        return redirect(url_for('auth'))
    return render_template('homepage.html', user=user)

@app.route('/settings', methods=['GET', 'POST'])
def settings():
    user = get_current_user()
    if not user:
        flash('Please log in to access settings.', 'error')
        return redirect(url_for('auth'))

    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        email = request.form.get('email', '').strip()
        display_name = request.form.get('display_name', '').strip()
        role = request.form.get('role', '').strip()
        tagline = request.form.get('tagline', '').strip()
        likes = request.form.get('likes', '').strip()
        dislikes = request.form.get('dislikes', '').strip()
        birthday = request.form.get('birthday', '').strip()
        place = request.form.get('place', '').strip()
        fav_games = request.form.get('fav_games', '').strip()

        if not username or not email:
            flash('Username and email are required.', 'error')
        elif not EMAIL_RE.match(email):
            flash('Please enter a valid email address.', 'error')
        elif username.lower() != user['username'].lower() and find_user(username) is not None:
            flash('The chosen username is already taken.', 'error')
        elif email.lower() != user['email'].lower() and find_user(email) is not None:
            flash('The chosen email is already taken.', 'error')
        else:
            update_user(user['username'],{
                'username': username,
                'email': email,
                'display_name': display_name,
                'role': role,
                'tagline': tagline,
                'likes': likes,
                'dislikes': dislikes,
                'birthday': birthday,
                'place': place,
                'fav_games': fav_games,
            })
            session['username'] = username
            flash('Your profile has been updated.', 'success')
            return redirect(url_for('settings'))

    return render_template('settings.html', user=user)

@app.route('/logout')
def logout():
    session.pop('username', None)
    flash('You have successfully logged out.', 'success')
    return redirect(url_for('auth'))


@app.route('/images/<path:filename>')
def image(filename):
    return send_from_directory(Path(__file__).parent / 'images', filename)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
