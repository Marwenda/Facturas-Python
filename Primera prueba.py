def main():
    print('hello python')
    num = (int(input('introduzca un numero')))
    print(num)

if __name__ == '__main__':
    import sys
    sys.exit(int(main()or 0))