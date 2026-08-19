import getpass

from werkzeug.security import generate_password_hash

if __name__ == "__main__":
    password = getpass.getpass("Enter the app password: ")
    if not password:
        raise SystemExit("password cannot be empty")
    print("\nAdd this to the server environment as PASSWORD_HASH:\n")
    print(generate_password_hash(password))