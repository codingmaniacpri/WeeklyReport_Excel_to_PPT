class Cell:
    def __init__(self, value, comment=None, formatting=None):
        self.value = value
        self.comment = comment
        self.formatting = formatting  # e.g., color, border

class Row:
    def __init__(self, cells):
        self.cells = cells  # List[Cell]

class Table:
    def __init__(self, rows):
        self.rows = rows  # List[Row]

class Sheet:
    def __init__(self, name, tables):
        self.name = name
        self.tables = tables  # List[Table]
