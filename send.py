import gmail

def main():
	"""Create a message for an email.

	Args:
		sender: Email address of the sender.
		to: Email address of the receiver.
		subject: The subject of the email message.
		message_text: The text of the email message.
		file: The path to the file to be attached.

	"""

	create_message_with_attachment('me', 'krupod@gmail.com', 'Test email with attachment', "Contents of the body", 'test-attachment.xls')


main()