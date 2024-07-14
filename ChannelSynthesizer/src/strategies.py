class ColumnStrategy:
    """
    Abstract base class for defining column strategies.
    Subclasses must implement the define_columns method.
    """

    def define_columns(self, page, num_columns):
        """
        Define the column boundaries for a given page and number of columns.

        Args:
            page (pdfplumber.Page): The PDF page to process.
            num_columns (int): Number of columns to split the page into.

        Returns:
            list of tuples: Each tuple contains the coordinates (x0, top, x1, bottom) for a column.
        """
        raise NotImplementedError("Subclasses must implement this method")


class TelenetColumnStrategy(ColumnStrategy):
    """
    Column strategy specific to the Telenet PDF template.
    """

    def define_columns(self, page, num_columns):
        buffers_start = [0] * num_columns
        buffers_end = [0] * num_columns

        if num_columns >= 7:
            buffers_end[4] = 12
            buffers_start[5] = 0
            buffers_end[5] = 11
            buffers_start[6] = -9

        return self.calculate_columns(page, num_columns, buffers_start, buffers_end)

    def calculate_columns(self, page, num_columns, buffers_start, buffers_end):
        """
        Calculates column boundaries with specified buffer adjustments.

        Args:
            page (pdfplumber.Page): The PDF page to process.
            num_columns (int): Number of columns to define.
            buffers_start (list): Start buffer for each column.
            buffers_end (list): End buffer for each column.

        Returns:
            list of tuples: Each tuple contains the coordinates (x0, top, x1, bottom) for a column.
        """
        column_width = page.width / num_columns

        columns = []
        for i in range(num_columns):
            x0 = i * column_width + buffers_start[i] if i != 0 else 0
            x1 = (i + 1) * column_width - buffers_end[i] if i != num_columns - 1 else page.width
            x0 = max(0, min(x0, page.width))
            x1 = max(x0, min(x1, page.width))
            columns.append((x0, 0, x1, page.height))
        return columns


class VOOColumnStrategy(ColumnStrategy):
    """
    Column strategy specific to the VOO PDF template.
    """


    def define_columns(self, page, num_columns):
        return self.calculate_columns(page, num_columns)

    def calculate_columns(self, page, num_columns, buffers_start=None, buffers_end=None):
        """
        Calculates column boundaries with optional buffer adjustments.

        Args:
            page (pdfplumber.Page): The PDF page to process.
            num_columns (int): Number of columns to define.
            buffers_start (list): Start buffer for each column.
            buffers_end (list): End buffer for each column.

        Returns:
            list of tuples: Each tuple contains the coordinates (x0, top, x1, bottom) for a column.
        """
        column_width = page.width / num_columns
        if buffers_start is None:
            buffers_start = [0] * num_columns
        if buffers_end is None:
            buffers_end = [0] * num_columns

        columns = []
        for i in range(num_columns):
            x0 = i * column_width + buffers_start[i] if i != 0 else 0
            x1 = (i + 1) * column_width - buffers_end[i] if i != num_columns - 1 else page.width
            x0 = max(0, min(x0, page.width))
            x1 = max(x0, min(x1, page.width))
            columns.append((x0, 0, x1, page.height))
        return columns
