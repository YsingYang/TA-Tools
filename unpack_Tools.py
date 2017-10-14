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
        self.dist_path = None # 作业放入的目的地址
        self._search_path = None # 搜索路径
        self._student_list = dict() # 学生列表 str->bool

    def create_dist_dir(self, path): # 传入一个路径, 创建相应的dir
        if(os.path.exists(path)):#检测路径文件是否存在
            print('该路径已经存在, 如果继续使用该路径输入yes,  否则程序终止')
            if(not input() == 'yes'):
                exit(0) #如果输入的不是yes则结束程序
            return # 如果是yes直接退出该函数
        self.dist_path = path # 同时保存作业地址
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

    def find_homework_result(self, path):
        self.set_search_dir(path)
        student_file_list = self._get_specific_file_list(True)
        student_list = self._get_sid(student_file_list)
        for student in student_list:
            #assert student in self._student_list

            self._student_list[student] = True
        result = [key for key, value in self._student_list.items() if value == False]
        return result



    def unpack_zip(self):
        pass

    def unpack_rar(self):
        pass

    def _get_all_specific_file(self):
        pass

    def _check_is_not_MainActivity(self, type, file_path):
        pass

    def _delete_dir(self, path, sub_file_list):
        pass


if __name__ == '__main__':
    tool = Tools()
    tool.init_student_list('/home/ysing/下载/手机平台应用开发.xls')
    print(tool.find_homework_result('./**/*.zip'))