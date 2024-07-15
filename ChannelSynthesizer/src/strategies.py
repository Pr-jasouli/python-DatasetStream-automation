class ColumnStrategy:
    def __init__(self, name="Configure strategies...", default_columns=2):
        self.name = name
        self.default_columns = default_columns
        self.max_columns = 8
        self.buffers_start = [0] * self.max_columns
        self.buffers_end = [0] * self.max_columns

    def update_name(self, new_name):
        self.name = new_name

    def update_buffers(self, buffers_start, buffers_end):
        if len(buffers_start) > self.max_columns or len(buffers_end) > self.max_columns:
            raise ValueError("Buffer lists exceed maximum column count")
        self.buffers_start = buffers_start
        self.buffers_end = buffers_end

    def define_columns(self, page, num_columns):
        raise NotImplementedError("Subclasses must implement this method")


def calculate_columns(page, num_columns, buffers_start, buffers_end):
    if num_columns <= 0:
        raise ValueError("Number of columns must be positive")
    column_width = page.width / num_columns
    columns = []

    for i in range(num_columns):
        start_buffer = buffers_start[i] if i < len(buffers_start) else 0
        end_buffer = buffers_end[i] if i < len(buffers_end) else 0

        x0 = max(0, i * column_width + start_buffer)
        x1 = min(page.width, (i + 1) * column_width - end_buffer) if i != num_columns - 1 else page.width

        columns.append((x0, 0, x1, page.height))

    return columns


class TelenetColumnStrategy(ColumnStrategy):
    def __init__(self, name="Telenet v1", default_columns=7):
        super().__init__(name, default_columns)
        self.update_buffers([0, 0, 0, 0, 0, 0, -9], [0, 0, 0, 0, 12, 11, 0])

    def define_columns(self, page, num_columns):
        return calculate_columns(page, num_columns, self.buffers_start, self.buffers_end)


class VOOColumnStrategy(ColumnStrategy):
    def __init__(self, name="VOO v1", default_columns=6):
        super().__init__(name, default_columns)
        self.update_buffers([15, -10, -10, -10, -5, -5, -5, -5], [19, 20, 10, 10, 5, 0, 0, 0])

    def define_columns(self, page, num_columns):
        return calculate_columns(page, num_columns, self.buffers_start, self.buffers_end)


class OrangeColumnStrategy(ColumnStrategy):
    def __init__(self, name="Orange v1", default_columns=3):
        super().__init__(name, default_columns)

    def define_columns(self, page, num_columns):
        return calculate_columns(page, num_columns, self.buffers_start, self.buffers_end)
