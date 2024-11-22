import random
import string
from flask import Blueprint, render_template, redirect, url_for, flash
from app import db
from app.models import Password
from app.forms import PasswordForm

main = Blueprint('main', __name__)

def generate_password(length=12):
    characters = string.ascii_letters + string.digits + string.punctuation
    return ''.join(random.choice(characters) for _ in range(length))

@main.route('/', methods=['GET', 'POST'])
def index():
    form = PasswordForm()
    if form.validate_on_submit():
        password = generate_password()
        new_password = Password(
            account=form.account.data,
            username=form.username.data,
            password=password
        )
        db.session.add(new_password)
        db.session.commit()
        flash(f'Password for {form.account.data} generated!', 'success')
        return redirect(url_for('main.index'))
    passwords = Password.query.all()
    return render_template('index.html', form=form, passwords=passwords)
