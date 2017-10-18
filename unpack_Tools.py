import shutil
import glob
import os
import rarfile
import re
import zipfile
import xlrd


class Tools:
    def __init__(self):
        self._dist_path = None  # 作业放入的目的地址
        self._search_path = None  # 搜索路径
        self._student_list = dict()  # 学生列表 str->bool

    def create_dist_dir(self, path):  # 传入一个路径, 创建相应的dir
        if(os.path.exists(path)):  # 检测路径文件是否存在
            print('该路径已经存在, 如果删除该路径输入yes,  否则继续')
            if(input() == 'yes'):
                shutil.rmtree(path)  # 删除掉该路径
        self._dist_path = path  # 同时保存作业地址
        try:
            os.makedirs(path)
        except:
            pass

    '''
    通过读取excel表格初始化学生数据
    path : excel表格路径
    col : 学号所在的列号
    start_row : 读取的起始行号
    '''
    # 初始化学生列表, 传入excel路径, 与学号所在的列, 和起始的行数
    def init_student_list(self, path, col=0, start_row=0):
        excel_data = xlrd.open_workbook(path)
        sh = excel_data.sheet_by_index(0)
        for row in range(start_row, sh.nrows):
            data = sh.cell(row, col).value
            self._student_list[data] = False


    # search_dir get and set
    def set_search_dir(self, path):
        self._search_path = path

    def get_search_dir(self):
        self._search_path

    '''
    返回检测作业上交情况
    paths : 搜索的文件夹目录, 不支持递归, 仅在文件夹内搜索
    '''
    def get_homework_result(self, path):  # 这里直接列出所有文件即可
        student_file_list = os.listdir(path)  # 列出该目录下所有文件
        student_list = self._get_sid(student_file_list)
        for student in student_list:
            self._student_list[student] = True
        result = [key for key, value in self._student_list.items() if value == False]
        return result

    '''
    解压后, 拷贝相应的后缀文件到学生文件根目录下,
    path : 解压后文件夹所在的根目录
    remove : 拷贝完成后是否删除文件夹
    kwargs-types : 需要拷贝的文件类型
    kwargs-exclude :  补需要拷贝的文件
    '''
    def copy_specific_type(self, path, remove=False, **kwargs):
        student_list = os.listdir(path)
        types = kwargs['types']
        exclude = kwargs['exclude']
        for student in student_list:
            student_dir = path + '/' + student
            for type in types:
                file_list = glob.glob(student_dir + '/**/' + type, recursive=True)  # 对file_list进行for循环, search/app/src/
                for file_path in file_list:
                    if (self._check_is_not_MainActivity(type, file_path, '/app/src', exclude = exclude)):  # 如果不是在app/src目录下的.java文件, continue
                        continue
                    try:
                        shutil.copy(file_path, student_dir)
                    except shutil.SameFileError:  # 如果目录一一致, 则跳过
                        pass

            if remove == True: #如果为remove = True, 删除文件夹
                self.delete_dir(student_dir)

    '''
    解压rar文件到指定目录
    path : 解压的目的目录
    types : 解压文件类型
    '''
    def unpack(self, path, types=['*.zip', '*.rar']):
        for type in types:
            assert self._dist_path != None  # 先设置好dist_path
            self._search_path = self._set_glob_search_file(path, type)
            files = self._get_specific_file_list(True)  # 获取所有rar列表
            regex_pattern = re.compile(r'\d{8}')  # 定义正则pattern
            for file in files:
                # 其实这里是否需要加入assert
                print(file + '    正在解压')
                sid = regex_pattern.findall(file)
                if (not len(sid) == 1):  # 如果获取到的sid不等于1
                    print('正则匹配出现异常, 该文件名为 {}'.format(file))
                    continue
                student_dir = self._dist_path + '/' + sid[0]
                try:
                    os.makedirs(student_dir)  # 创建多级目录
                except FileExistsError:  # 文件已经存在
                    pass
                try:
                    compressed_object = rarfile.RarFile(file) if file.endswith('.rar') else zipfile.ZipFile(file) # 可拓展性不强
                    compressed_object.extractall(student_dir)
                except:  # 避免一些错误的压缩文件
                    pass

    def delete_dir(self, path):
        sub_file_list = os.listdir(path)
        for sub_file in sub_file_list:  # 注意列出来的只是一个文件名, 而不是一个目录
            if os.path.isdir(os.path.join(path, sub_file)):  # 如果是目录, 则删除,
                shutil.rmtree(os.path.join(path, sub_file))

    '''
    设置glob搜索路径
    path : 搜索根文件
    file_type : 搜索的文件类型
    '''
    def _set_glob_search_file(self, path, file_type):
        return os.path.join(path, '**', file_type) # 方便跨平台

    # 文件搜索, 参数recursive表示是否递归
    def _get_specific_file_list(self, recursive): #获取文件列表
        return glob.glob(self._search_path, recursive=recursive)  # 是否递归

    def _check_is_not_MainActivity(self, type, file_path, specific_dir = None, exclude = []): # 写得有点问题
        if(specific_dir != None and type == '*.java'):
            if (re.search(specific_dir, file_path) == None):
                return True  # 如果不存在app/src, 则输出文件返回True
        for exclude_postfix in exclude: # 如果是exclude中的文件则跳过, 比如android的test文件
            if re.search(exclude_postfix, file_path) != None:
                return True
        return False

    '''
    获取学生学号列表
    '''
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