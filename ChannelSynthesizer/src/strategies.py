class ColumnStrategy:
    def __init__(self, name="Configure strategies..."):
        self.name = name
        self.max_columns = 8
        self.buffers_start = [0] * self.max_columns
        self.buffers_end = [0] * self.max_columns

    def update_name(self, new_name):
        self.name = new_name

    def update_buffers(self, buffers_start, buffers_end):
        self.buffers_start = buffers_start[:self.max_columns]
        self.buffers_end = buffers_end[:self.max_columns]

    def define_columns(self, page, num_columns):
        raise NotImplementedError("Subclasses must implement this method")


class TelenetColumnStrategy(ColumnStrategy):
    def __init__(self, name="Telenet Original"):
        super().__init__(name)
        if self.max_columns >= 7:
            self.buffers_end[4] = 12
            self.buffers_start[5] = 0
            self.buffers_end[5] = 11
            self.buffers_start[6] = -9

    def define_columns(self, page, num_columns):
        return self.calculate_columns(page, num_columns, self.buffers_start, self.buffers_end)

    def calculate_columns(self, page, num_columns, buffers_start, buffers_end):
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
    def __init__(self, name="VOO Original"):
        super().__init__(name)


    def define_columns(self, page, num_columns):
        return self.calculate_columns(page, num_columns, self.buffers_start, self.buffers_end)

    def calculate_columns(self, page, num_columns, buffers_start, buffers_end):
        column_width = page.width / num_columns
        columns = []
        for i in range(num_columns):
            x0 = i * column_width + buffers_start[i] if i != 0 else 0
            x1 = (i + 1) * column_width - buffers_end[i] if i != num_columns - 1 else page.width
            x0 = max(0, min(x0, page.width))
            x1 = max(x0, min(x1, page.width))
            columns.append((x0, 0, x1, page.height))
        return columns
