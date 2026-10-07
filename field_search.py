def filter_fields(all_fields, search_text):
    """
    Filters a list of field names based on the search text.
    It performs a case-insensitive search.
    """
    if not search_text:
        return all_fields
        
    search_text = search_text.lower()
    return [field for field in all_fields if search_text in str(field).lower()]
