import pandas as pd

import glob


#파일 Union
all_data = pd.DataFrame()
for f in glob.glob('C:/Users/LXPER MINI001/Desktop/잡업/temp/고1,2/*.xlsx'): # 예를들어 201901, 201902 로 된 파일이면 2019_*

    df = pd.read_excel(f)
    try:
        all_data = all_data.append(df, ignore_index=True)
        print(f, "done")
    except:
        print(f, "failed")
#데이터갯수확인
print(all_data.shape)

#데이터 잘 들어오는지 확인
all_data.head()

#파일저장
all_data.to_excel("C:/Users/LXPER MINI001/Desktop/잡업/고1,고2/고1,2MERGED.xlsx", header=False, index=False)