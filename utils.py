import re

def clean_date_string(date_str):
    """
    Remove timezone information from a date string.
    Example: '10/14/2024 11:47:53 AM ET' -> '10/14/2024 11:47:53 AM'
    """
    if date_str:
        return re.sub(r'\s*\b(ET|EST|EDT)\b$', '', date_str.strip())
    else:
        return ''
