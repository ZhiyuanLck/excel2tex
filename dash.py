class T:
    def __init__(self, width, mode):
        self.n = int(width / 0.1)
        self.mode = mode

    def __iter__(self):
        self.iter = 1
        return self

    def __next__(self):
        if self.iter <= self.n:
            self.iter += 1
            return '0.1pt 0pt' if self.mode == 'dash' else '0pt 0.1pt'
        else:
            raise StopIteration

if __name__ == '__main__':
    a = float(input('dash: '))
    b = float(input('gap: '))
    res = ''
    for s in iter(T(a, 'dash')):
        res += s + ' '
    for s in iter(T(a, 'gap')):
        res += s + ' '
    res = res.strip()
    print(res)
