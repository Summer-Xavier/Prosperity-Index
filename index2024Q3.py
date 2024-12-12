import pandas as pd


class IndexQ3:

    def __init__(self,file):
        self.file = file
        self.cleaned_data = None
        self.firm_score = None


    def clean_data(self):
        df = pd.read_excel(self.file,sheet_name='真')
        df.drop_duplicates(subset='单位详细名称:', inplace=True)
        df.set_index('单位详细名称:',inplace=True)
        df = df.iloc[:,11:19]
        rep_dic = {r'明显(增加|上升)':'++',
                   r'有所(增加|上升)':'+',
                   r'基本不变':'0',
                   r'有所(减少|下降)':'-',
                   r'明显(减少|下降)':'--'}
        df.replace(rep_dic,regex=True,inplace=True)
        self.cleaned_data = df


    def calculate_index(self,df):
        ncols = df.shape[1]
        df0 = df.iloc[:,0].value_counts(normalize=True)
        df0.name = df.columns[0]
        for i in range(1,ncols):
            df1 = df.iloc[:,i].value_counts(normalize=True)
            df1.name = df.columns[i]
            df0 = pd.concat([df0,df1],axis=1)

        df0 = df0.T
        df0['score'] = (df0['++']
                        + 0.5 * df0['+']
                        - 0.5 * df0['-']
                        - df0['--']) * 50 + 50

        return df0


    def get_firms_scores(self):
        rep_dic = {'++':5,
                   '+':4,
                   '0':3,
                   '-':2,
                   '--':1}
        df = self.cleaned_data.replace(rep_dic)
        # Key part : weights for 8 indices
        weights = [0,
                   1,
                   0,
                   0,
                   0,
                   1,
                   0,
                   0]
        df['firm_score'] = (df.iloc[:,0] * weights[0] +
                            df.iloc[:,1] * weights[1] +
                            df.iloc[:,2] * weights[2] +
                            df.iloc[:,3] * weights[3] +
                            df.iloc[:,4] * weights[4] +
                            df.iloc[:,5] * weights[5] +
                            df.iloc[:,6] * weights[6] +
                            df.iloc[:,7] * weights[7])

        self.firm_score = df['firm_score']


    def get_it_right(self,start=0,end=200):
        df = pd.concat([self.cleaned_data, self.firm_score], axis=1)
        df.sort_values(by='firm_score',inplace=True)
        df = df.iloc[start:end,0:8]
        res = self.calculate_index(df)
        return [res, df]





if __name__ == '__main__':
    N = 200
    start = 46
    end = start + N
    outfile = r'result.csv'
    outfile2 = r'data.csv'

    path = r'C:\Users\XF\Desktop\index_q4'
    fileName = '新的.xlsx'
    i = IndexQ3(path + '\\' + fileName)
    i.clean_data()
    i.get_firms_scores()
    res, firm_rank = i.get_it_right(start,end)
    res.to_csv(path+'\\'+outfile,index=True,encoding='gbk')
    firm_rank.to_csv(path+'\\'+outfile2,index=True,encoding='gbk')
    print(res['score'])
