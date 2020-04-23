def scan_all(cell_list, check):
    '''
    return true if all checked false
    '''
    for cell in cell_list:
        if getattr(cell, check)():
            return False
    return True

# wrap !{}
def wrap_excl(text):
    return f'  !{{{text}}}\n'

# wrap >{}
def wrap_ge(text):
    return f'  >{{{text}}}\n'
