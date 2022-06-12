def JoinText(headers:list) -> str:
	"""Join strings by change line code.
	
	Join strings by change line code as headers

	Args:
		headetrs(list): Collection of strings to set into header.
						One item is one line in header.
						When output, all items are joined by change line code.

	Returns:
		String to be set into header.
	"""
	header_text = ''
	is_top = True
	for header_item in headers:
		if False == is_top:
			header_text += '\n'
		header_text += header_item
		is_top = False

	return header_text

