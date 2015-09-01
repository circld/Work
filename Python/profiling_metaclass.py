"""
This is an IronPython script; will not run with CPython
"""

from System.Diagnostics import Stopwatch


class Fib(object):

    def __call__(self, n):
        self.fib_list = list()
        self.__calc_fib(n)
        return self.fib_list

    def __calc_fib(self, n):
        for idx in xrange(n):
            if idx == 0:
                self.fib_list.append(0)
            elif idx == 1:
                self.fib_list.append(1)
            else:
                self.fib_list.append(
                    self.fib_list[-2] + self.fib_list[-1]
                )


def F():
    a,b = 0,1
    yield a
    yield b
    while True:
        a, b = b, a + b
        yield b


if __name__ == '__main__':

    sw = Stopwatch()

    fib_calc = Fib()
    f = F()

    sw.Start()
    print fib_calc(18)
    sw.Stop()

    time_taken = sw.ElapsedMilliseconds
    print time_taken

    sw.Start()
    print [f.next() for i in xrange(18)]
    sw.Stop()

    print sw.ElapsedMilliseconds
