import random

def generate_random_key(length=64):
    chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789@#$%^&*()"
    key = ""
    for i in range(length):
        key += random.choice(chars)
    return key

# Generate a random key of length 16
key = generate_random_key()

# Print the key
print(key)
