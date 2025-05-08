from config import db, logger
from werkzeug.security import generate_password_hash, check_password_hash
from pymongo.errors import DuplicateKeyError

def init_users_collection():
    """Initialize the users collection with a unique index on username if it doesn't exist."""
    collection = db["users"]
    index_name = "username_idx"
    
    # Check existing indexes
    existing_indexes = collection.index_information()
    
    # Check if an index on 'username' field already exists (regardless of name)
    username_index_exists = any(
        index.get('key', []) == [('username', 1)] or index.get('key', []) == [('username', -1)]
        for index in existing_indexes.values()
    )
    
    if not username_index_exists:
        try:
            collection.create_index("username", unique=True, name=index_name)
            logger.info("Created unique index on username for users collection")
        except Exception as e:
            logger.error(f"Failed to create index for users collection: {str(e)}")
            raise
    else:
        logger.info("Unique index on username already exists, skipping creation")

def create_user(username, email, password):
    """Create a new user with hashed password."""
    collection = db["users"]
    hashed_password = generate_password_hash(password, method='pbkdf2:sha256')
    user = {
        "username": username,
        "email": email,
        "password": hashed_password,
        "role": "user",  # Default role
        "blocked": False
    }
    try:
        collection.insert_one(user)
        logger.info(f"User {username} created successfully")
        return True
    except DuplicateKeyError:
        logger.warning(f"Username {username} already exists")
        return False
    except Exception as e:
        logger.error(f"Failed to create user {username}: {str(e)}")
        return False


def update_user_role(username, new_role):
    """Change a user's role (e.g. user → admin)."""
    result = db["users"].update_one(
        {"username": username},
        {"$set": {"role": new_role}}
    )
    return result.modified_count == 1

def block_user(username):
    """Mark a user as blocked (cannot log in)."""
    result = db["users"].update_one(
        {"username": username},
        {"$set": {"blocked": True}}
    )
    return result.modified_count == 1

def unblock_user(username):
    """Unblock a previously blocked user."""
    result = db["users"].update_one(
        {"username": username},
        {"$set": {"blocked": False}}
    )
    return result.modified_count == 1

def delete_user(username):
    """Permanently delete a user account."""
    result = db["users"].delete_one({"username": username})
    return result.deleted_count == 1

def verify_user(username, password):
    """Verify user credentials and return user data if valid."""
    collection = db["users"]
    user = collection.find_one({"username": username})
    if user and check_password_hash(user["password"], password):
        logger.info(f"User {username} verified successfully")
        return user
    logger.warning(f"Verification failed for user {username}")
    return None

def create_admin_if_not_exists(admin_username, admin_password):
    """Create admin user from environment variables if it doesn't exist."""
    collection = db["users"]
    if not collection.find_one({"username": admin_username}):
        hashed_password = generate_password_hash(admin_password, method='pbkdf2:sha256')
        user = {
            "username": admin_username,
            "email": "admin@example.com",  # Placeholder email
            "password": hashed_password,
            "role": "admin"
        }
        try:
            collection.insert_one(user)
            logger.info(f"Admin user {admin_username} created successfully")
        except Exception as e:
            logger.error(f"Failed to create admin user {admin_username}: {str(e)}")
            raise
