from flask import Flask, render_template, request, redirect, url_for, session, flash
import pandas as pd
import random
import pyttsx3
import speech_recognition as sr

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'  # Needed for session management

excel_file = 'bank.xlsx'

# Initialize text-to-speech engine
action = pyttsx3.init()

def text_to_speech(message):
    action.say(message)
    action.runAndWait()
    action.stop()

# Helper functions (same as your original code with minor adjustments)
def reset_password(user_id, new_pass):
    contact = pd.read_excel(excel_file, sheet_name="Contact")
    if user_id in contact['Name'].values:
        # Update password in the DataFrame
        contact.loc[contact['Name'] == user_id, 'Password'] = new_pass
        # Save back to Excel
        with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            contact.to_excel(writer, sheet_name="Contact", index=False)
        return True
    return False

def create_account(new_user, new_pass):
    contact = pd.read_excel(excel_file, sheet_name="Contact")
    balance = pd.read_excel(excel_file, sheet_name="Balance")
    
    if new_user in contact['Name'].values:
        return False  # User already exists
    
    # Add new user to contact
    new_contact = pd.DataFrame({'Name': [new_user], 'Password': [new_pass]})
    contact = pd.concat([contact, new_contact], ignore_index=True)
    
    # Add new user balance
    ran = random.randint(1000, 909909)
    new_balance = pd.DataFrame({'Name': [new_user], 'Balance': [ran]})
    balance = pd.concat([balance, new_balance], ignore_index=True)
    
    # Save both sheets
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        contact.to_excel(writer, sheet_name='Contact', index=False)
        balance.to_excel(writer, sheet_name='Balance', index=False)
    
    return True

def get_balance(username):
    balance_df = pd.read_excel(excel_file, sheet_name='Balance')
    user_balance = balance_df.loc[balance_df['Name'] == username, 'Balance'].values
    return user_balance[0] if len(user_balance) > 0 else None

def update_balance(username, amount, operation='deposit'):
    balance_df = pd.read_excel(excel_file, sheet_name='Balance')
    if username in balance_df['Name'].values:
        if operation == 'deposit':
            balance_df.loc[balance_df['Name'] == username, 'Balance'] += amount
        elif operation == 'withdraw':
            balance_df.loc[balance_df['Name'] == username, 'Balance'] -= amount
        
        with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            balance_df.to_excel(writer, sheet_name='Balance', index=False)
        return True
    return False

# Routes
@app.route('/')
def home():
    return render_template('index.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        
        try:
            password = int(password)
        except ValueError:
            flash('Password must be numeric', 'error')
            return redirect(url_for('login'))
        
        contact_df = pd.read_excel(excel_file, sheet_name='Contact')
        contact = dict(zip(contact_df['Name'], contact_df['Password']))
        
        if username in contact and contact[username] == password:
            session['username'] = username
            return redirect(url_for('dashboard'))
        else:
            flash('Invalid username or password', 'error')
    
    return render_template('login.html')

@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        confirm_password = request.form['confirm_password']
        
        if password != confirm_password:
            flash('Passwords do not match', 'error')
            return redirect(url_for('register'))
        
        try:
            password = int(password)
        except ValueError:
            flash('Password must be numeric', 'error')
            return redirect(url_for('register'))
        
        if create_account(username, password):
            flash('Account created successfully! Please login.', 'success')
            return redirect(url_for('login'))
        else:
            flash('Username already exists', 'error')
    
    return render_template('register.html')

@app.route('/reset-password', methods=['GET', 'POST'])
def reset_password_page():
    if request.method == 'POST':
        username = request.form['username']
        new_pass = request.form['new_password']
        confirm_pass = request.form['confirm_password']
        
        if new_pass != confirm_pass:
            flash('Passwords do not match', 'error')
            return redirect(url_for('reset_password_page'))
        
        try:
            new_pass = int(new_pass)
        except ValueError:
            flash('Password must be numeric', 'error')
            return redirect(url_for('reset_password_page'))
        
        if reset_password(username, new_pass):
            flash('Password reset successfully! Please login.', 'success')
            return redirect(url_for('login'))
        else:
            flash('Username not found', 'error')
    
    return render_template('reset_password.html')

@app.route('/dashboard')
def dashboard():
    if 'username' not in session:
        return redirect(url_for('login'))
    
    username = session['username']
    balance = get_balance(username)
    
    return render_template('dashboard.html', username=username, balance=balance)

@app.route('/deposit', methods=['POST'])
def deposit():
    if 'username' not in session:
        return redirect(url_for('login'))
    
    username = session['username']
    try:
        amount = float(request.form['amount'])
        if amount <= 0:
            flash('Amount must be positive', 'error')
        else:
            if update_balance(username, amount, 'deposit'):
                flash(f'Successfully deposited ${amount:.2f}', 'success')
            else:
                flash('Transaction failed', 'error')
    except ValueError:
        flash('Invalid amount', 'error')
    
    return redirect(url_for('dashboard'))

@app.route('/withdraw', methods=['POST'])
def withdraw():
    if 'username' not in session:
        return redirect(url_for('login'))
    
    username = session['username']
    try:
        amount = float(request.form['amount'])
        balance = get_balance(username)
        
        if amount <= 0:
            flash('Amount must be positive', 'error')
        elif amount > balance:
            flash('Insufficient funds', 'error')
        else:
            if update_balance(username, amount, 'withdraw'):
                flash(f'Successfully withdrew ${amount:.2f}', 'success')
            else:
                flash('Transaction failed', 'error')
    except ValueError:
        flash('Invalid amount', 'error')
    
    return redirect(url_for('dashboard'))

@app.route('/logout')
def logout():
    session.pop('username', None)
    return redirect(url_for('home'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)