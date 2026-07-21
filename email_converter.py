import os
import re
import extract_msg

email_path = os.path.join(os.getcwd(), 'Email')
output_dir = os.path.join(os.getcwd(), 'Txt')

def convert_msg_to_txt(msg_file, output_dir):
    """

    This function is used to convert .msg files from outlook
    into .txt files
    """
    # Open the .msg file
    msg = extract_msg.Message(msg_file)
    # Extract the content of the message
    subject = msg.subject
    body = msg.body
    to = msg.to
    # Clean the body from Content ID related to possible images
    body_cleaned = re.sub(r'\[cid:.*?]', "", body)
    # Extract the name of the file
    filename = os.path.splitext(os.path.basename(msg_file))[0]
    # Define the .txt filename
    new_name = "".join([c if c.isalnum() else " " for c in filename])
    # Output path
    txt_filename = os.path.join(output_dir, f"{new_name}.txt")
    # Save the file as a .txt
    with open(txt_filename, 'w', encoding="utf-8") as txt_file:
        txt_file.write(to)
        txt_file.write(f"{subject}\n\n")
        txt_file.write(body_cleaned)

    print(f"Converted: {msg} -> {txt_filename}")

# Iterate over the folder
for file in os.listdir(email_path):
    if file.endswith(".msg"):
        convert_msg_to_txt(os.path.join(email_path, file), output_dir)
    else:
        continue




