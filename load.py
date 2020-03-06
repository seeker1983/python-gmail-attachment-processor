import gmail

def main():
  """ Example of loading email attachments
  
  Args:
	folder: Location to store files.
  """	
  
  gmail.load_unread('files', ['krupod@gmail.com'], "label_excel_loaded_2")



main()