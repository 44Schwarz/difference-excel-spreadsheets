import pandas as pd


def compare(file1, file2):
    try:
        df = pd.read_excel(file1, header=None)
        df2 = pd.read_excel(file2, header=None)
    except FileNotFoundError:
        print("One of the files doesn't exist")
        return

    if min(df.shape[0], df2.shape[0]) == 0:
        print("One of the files is empty")
        return

    df.fillna('', inplace=True)
    df2.fillna('', inplace=True)

    differences = []

    for j in range(min(df.shape[1], df2.shape[1])):
        for i in range(min(df.shape[0], df2.shape[0])):
            if df[j][i] != df2[j][i]:
                # print('{} | {}'.format(df[j][i], df2[j][i]))
                differences.append(((i + 1, j + 1), df[j][i], df2[j][i]))

        for i in range(min(df.shape[0], df2.shape[0]), max(df.shape[0], df2.shape[0])):
            if df.shape[0] > df2.shape[0]:
                if df[j][i] != '':
                    # print('{} | '.format(df[j][i]))
                    differences.append(((i + 1, j + 1), df[j][i], ''))
            if df.shape[0] < df2.shape[0]:
                if df2[j][i] != '':
                    # print(' | {}'.format(df2[j][i]))
                    differences.append(((i + 1, j + 1), '', df2[j][i]))

    for j in range(min(df.shape[1], df2.shape[1]), max(df.shape[1], df2.shape[1])):
        for i in range(min(df.shape[0], df2.shape[0])):
            if df.shape[1] > df2.shape[1]:
                if df[j][i] != '':
                    # print('{} | '.format(df[j][i]))
                    differences.append(((i + 1, j + 1), df[j][i], ''))
            if df.shape[1] < df2.shape[1]:
                if df2[j][i] != '':
                    # print(' | {}'.format(df2[j][i]))
                    differences.append(((i + 1, j + 1), '', df2[j][i]))

    # print(df.shape[0], df.shape[1])
    # print(differences)
    df_for_save = pd.DataFrame({"Cell": list(zip(*differences))[0], "First file": list(zip(*differences))[1],
                                "Second file": list(zip(*differences))[2]})
    # print(df_for_save)
    df_for_save.to_excel(r'Result_comparison.xlsx', index=False)


def main():
    file1 = input("Path to the first file: ")
    file2 = input("Path to the second file: ")

    compare(file1, file2)


if __name__ == "__main__":
    main()



