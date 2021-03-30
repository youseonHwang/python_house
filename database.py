import pandas as pd


def save(location, cleaness, built_in):  # 저장하는 함수
    idx = len(pd.read_csv("database.csv"))  # database.csv파일의 길이 읽어들이기
    new_df = pd.DataFrame({"location": location,  # 새로운 dataFrame 생성
                           "cleaness": cleaness,
                           "built_in": built_in}, index=[idx])
    new_df.to_csv("database.csv", mode="a", header=False)
    return None


def load_list():  # 리스트를 가져오는 함수
    house_list = []
    df = pd.read_csv("database.csv")
    for i in range(len(df)):
        house_list.append(df.iloc[i].tolist())  # iloc은 row선택
    return house_list


def now_index():  # 현재 인덱스번호 가져오기
    df = pd.read_csv("database.csv")
    return len(df)-1


def load_house(idx):
    df = pd.read_csv("database.csv")
    house_info = df.iloc[idx]  # parameter로 받아온 idx값으로 row선택
    return house_info


if __name__ == "__main__":
    load_list()