import re
from O365.excel import Range as _Range

class Range(_Range):
    pattern = r'^.*!(?P<left>[A-Z]+)(?P<top>[0-9]+)(:(?P<right>[A-Z]+)(?P<bottom>[0-9]+))?$'
    def __init__(self, address: str):
        super().__init__(address)
        self.matchgroup = re.search(self.pattern, self.address).groupdict()

    def update(self, values: list[list]):
        self.values = values
        super().update()

    @property
    def left(self):
        return self.matchgroup.get('left')

    @property
    def right(self):
        return self.matchgroup.get('right')

    @property
    def top(self):
        return self.matchgroup.get('top')

    @property
    def bottom(self):
        return self.matchgroup.get('bottom')