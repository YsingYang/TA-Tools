import shutil
import glob
import os
import re
import zipfile
import xlrd
'''
def check_is_not_MainActivity(type, file_path):
    if(type == types[1] and re.search('/app/src', file_path) == None):
        return True #如果不存在app/src, 则输出文件返回True
    if(re.search('ExampleUnitTest.java', file_path) != None or re.search('ExampleInstrumentedTest.java', file_path) != None):
        return True #如果是test文件夹中的内容返回true
    return False

def delete_dir(path, sub_file_list):
    for sub_file in sub_file_list: #注意列出来的只是一个文件名, 而不是一个目录
        if os.path.isdir(path + '/' + sub_file): #如果是目录, 则删除,
            shutil.rmtree(path + '/' + sub_file)

zip_list = glob.glob('./**/*.zip', recursive=True)
regex_pattern = re.compile(r'\d{8}')
for zip_file in zip_list:
    #首先创建相应学号的文件夹
    sid = regex_pattern.findall(zip_file)
    if(not len(sid) == 1): #如果获取到的sid不等于1
        print('正则匹配出现异常, 该文件名为 {}'.format(zip_file))
        continue

    # 指定文件中创建目录
    student_dir = '../Lab_4/' + sid[0]
    try:
        os.makedirs(student_dir) #创建多级目录
    except FileExistsError:
        pass
    #将文件解压至文件夹中
    #zip_object = zipfile.ZipFile(zip_file)
    #zip_object.extractall(student_dir)

    # 搜索相应学号文件夹的.java, .pdf文件并放入相应学号文件夹中
    types = ['*.pdf', '*.java']
    for type in types:
        file_list = glob.glob(student_dir + '/**/' + type, recursive=True) # 对file_list进行for循环, search/app/src/
        for file_path in file_list:
            if(check_is_not_MainActivity(type, file_path)): #如果不是在app/src目录下的.java文件, continue
                continue
            try:
                shutil.copy(file_path, student_dir)
            except shutil.SameFileError: #如果目录一一致, 则跳过
                pass
    #对student_dir下进行list, 如果有
    sub_file_list = os.listdir(student_dir)
    delete_dir(student_dir, sub_file_list)
'''





class Tools:
    def __init__(self):
        self.base_path = None # 存放压缩文件的基本路径
        self._dist_path = None # 作业放入的目的地址
        self._search_path = None # 搜索路径
        self._student_list = dict() # 学生列表 str->bool

    def create_dist_dir(self, path): # 传入一个路径, 创建相应的dir
        if(os.path.exists(path)):#检测路径文件是否存在
            print('该路径已经存在, 如果继续使用该路径输入yes,  否则程序终止')
            if(not input() == 'yes'):
                exit(0) #如果输入的不是yes则结束程序
            #如果是Yes, 移除之中的内容, 并重新创建
            shutil.rmtree(path) #删除掉该路径
        self._dist_path = path # 同时保存作业地址
        os.makedirs(path)

    # search_dir get and set
    def set_search_dir(self, path):
        self._search_path = path

    def get_search_dir(self):
        self._search_path

    def init_student_list(self, path): # 初始化学生列表
        excel_data = xlrd.open_workbook(path)
        sh = excel_data.sheet_by_index(0)
        for row in range(6, sh.nrows):
            data = sh.cell(row, 1).value
            self._student_list[data] = False

    def _get_specific_file_list(self, recursive): #获取文件列表
        return glob.glob(self._search_path, recursive=recursive)  # 是否递归

    def _get_sid(self, file_list):
        regex_pattern = re.compile(r'\d{8}')
        result = list()
        for zip_file in file_list:
            # 首先创建相应学号的文件夹
            sid = regex_pattern.findall(zip_file)
            if (not len(sid) == 1):  # 如果获取到的sid不等于1
                print('正则匹配出现异常, 该文件名为 {}'.format(zip_file))
                continue
            result.append(sid[0])
        return result

    def get_homework_result(self, path):
        self.set_search_dir(path)
        student_file_list = self._get_specific_file_list(True)
        student_list = self._get_sid(student_file_list)
        for student in student_list:
            #assert student in self._student_list

            self._student_list[student] = True
        result = [key for key, value in self._student_list.items() if value == False]
        return result



    def unpack_zip(self, path):
        self.set_search_dir(path)
        assert self._dist_path != None # 先设置好dist_path
        homework_list = self._get_specific_file_list(True) #获取所有zip列表
        print(homework_list)
        regex_pattern = re.compile(r'\d{8}') #定义正则pattern
        for zip_file in homework_list:
            # 首先创建相应学号的文件夹
            sid = regex_pattern.findall(zip_file)
            if (not len(sid) == 1):  # 如果获取到的sid不等于1
                print('正则匹配出现异常, 该文件名为 {}'.format(zip_file))
                continue

            # 指定文件中创建目录
            student_dir = self._dist_path + '/' + sid[0]
            try:
                os.makedirs(student_dir)  # 创建多级目录
            except FileExistsError: # 文件已经存在
                pass
            # 将文件解压至文件夹中
            zip_object = zipfile.ZipFile(zip_file)
            zip_object.extractall(student_dir)


    def copy_specific_type(self, path, remove=False, **kwargs): # 传入解压后, 作业放入的base_dir

        student_list = os.listdir(path)
        types = kwargs['types']
        exclude = kwargs['exclude']
        for student in student_list:
            student_dir = path + '/' + student
            for type in types:
                file_list = glob.glob(student_dir + '/**/' + type, recursive=True)  # 对file_list进行for循环, search/app/src/
                print(file_list)
                for file_path in file_list:
                    if (self._check_is_not_MainActivity(type, file_path, '/app/src', exclude = exclude)):  # 如果不是在app/src目录下的.java文件, continue
                        continue
                    try:
                        shutil.copy(file_path, student_dir)
                    except shutil.SameFileError:  # 如果目录一一致, 则跳过
                        pass

            if remove == True: #如果为remove = True, 删除文件夹
                self.delete_dir(student_dir)


    def unpack_rar(self):
        pass

    def _get_all_specific_file(self):
        pass

    def _check_is_not_MainActivity(self, type, file_path, specific_dir = None, exclude = []): # 写得有点问题
        if(specific_dir != None and type == '*.java'):
            if (re.search(specific_dir, file_path) == None):
                return True  # 如果不存在app/src, 则输出文件返回True
        for exclude_postfix in exclude: # 如果是exclude中的文件则跳过, 比如android的test文件
            if re.search(exclude_postfix, file_path) != None:
                return True
        #if (re.search('ExampleUnitTest.java', file_path) != None or re.search('ExampleInstrumentedTest.java',
        #                                                                      file_path) != None):
        return False

    def delete_dir(self, path):
        sub_file_list = os.listdir(path)
        for sub_file in sub_file_list:  # 注意列出来的只是一个文件名, 而不是一个目录
            if os.path.isdir(path + '/' + sub_file):  # 如果是目录, 则删除,
                shutil.rmtree(path + '/' + sub_file)



if __name__ == '__main__':
    tool = Tools()
    tool.init_student_list('/home/ysing/下载/手机平台应用开发.xls')
    #tool.create_dist_dir('../Lab_4')
    #tool.unpack_zip('./**/*.zip') #搜索zip路径
    #tool.copy_specific_type('../Lab_4', remove=True, types = ['*.pdf', '*.java'], exclude=['ExampleUnitTest.java', 'ExampleInstrumentedTest.java'])

