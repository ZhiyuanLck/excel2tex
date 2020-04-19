def scan_all(cell_list, check):
    '''
    return true if all checked false
    '''
    for cell in cell_list:
        if getattr(cell, check)():
            return False
    return True
