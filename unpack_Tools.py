import shutil
import imaplib
import glob
import os
import copy
import re
import zipfile
import email
from email.header import decode_header
import sys
import xlrd
import xlwt
import xlutils.copy
import rarfile

class IMAP_Tools:
    def __init__(self):
        self._conn_imap_server = None

    '''
    设置imap服务器
    '''
    def set_imaplib_server(self, imap_server):
        self._conn_imap_server = imaplib.IMAP4_SSL(imap_server)

    def set_imap_select(self, select='INBOX'):  # 默认选取收信件
        try:
            result_code, messages = self._conn_imap_server.select(select)
            if result_code != 'OK':  # 选取的文件夹不存在
                print("selected mailbox doesn't exist")
                sys.exit(1)

        except:
            print('set_imap_select raise exception')

    '''
    通过账号密码登录imap服务器
    '''
    def login_imap(self, username, password):
        try:
            return self._conn_imap_server.login(username, password)
        except:
            print(sys.exc_info()[1])
            sys.exit(1)  # 连接异常直接退出

    def search_email(self, type='UNSEEN'):
        (result_code, messages) = self._conn_imap_server.search(None, type)
        if(result_code != "OK"):
            print('search_email error')
            sys.exit(1)
        return result_code, messages

    def download_from_emails(self, messages, base_dir=os.getcwd()):
        email_list = self._messages_to_list(messages)
        for email_object in email_list:
            response, data = self._conn_imap_server.fetch(email_object, "(RFC822)")   # 获取所有list
            # 这里不能直接用str()强制转换, 因为data[0][1]是byte类型, 需要进行decode
            message = email.message_from_bytes(data[0][1])
            # 检查是否有附件
            if message.get_content_maintype() != 'multipart':  # 不为multipart对象
                continue
            seen_flag = False  # 标识已读
            for part in message.walk():
                # just multipart container
                if part.get_content_maintype() == 'multipart':
                    print(part.as_string())
                    continue
                # 参考StackOverflow的
                if part.get('Content-Disposition') is None:
                    print(part.as_string())
                    continue
                filename = part.get_filename()
                #print(decode_header(filename)[0])
                filename_decoded = decode_header(filename)[0]  # tuple
                if(filename_decoded[1] != None and filename_decoded[1] != 'NoneType'):
                    filename = filename_decoded[0].decode(filename_decoded[1])  # 可能存在utf-8编码
                print('准备下载 : ' + filename)
                attach_path = os.path.join(base_dir, filename)
                if not filename.endswith('.zip') and not filename.endswith('.rar'):
                    continue
                if not os.path.isfile(attach_path):
                    with open(attach_path, 'wb') as fp:
                        fp.write(part.get_payload(decode=True))
                        seen_flag = True
            if seen_flag:
                self._conn_imap_server.store(email_object, '+FLAGS', '\Seen')  # 标记邮件为已读

    def imap_logout(self):
        self._conn_imap_server.logout()

    def _messages_to_list(self, messages):
        return messages[0].split()


class Tools:
    def __init__(self):
        self._dist_path = None  # 作业放入的目的地址
        self.search_path = None  # 搜索路径
        self._student_list = dict()  # 学生列表 str->bool

    @property
    def search_path(self):
        self._search_path

    @search_path.setter
    def search_path(self, path):
        self._search_path = path

    '''
    创建存在作业目录的dir, 相当于解压的目的地址
    path : 文件夹目的路径
    如果该路径存在 则会阻塞在输入状态。 如果输入是yes(不区分大小写), 则删除该文件夹
    '''
    def create_dist_dir(self, path):  # 传入一个路径, 创建相应的dir
        if os.path.exists(path):  # 检测路径文件是否存在
            print('该路径已经存在, 如果删除该路径输入yes,  否则继续')
            if input().lower() == 'yes':
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
    根据homework_result来编辑作业提交情况的excel表格
    path : excel文件的路径
    sid_col : 读取excel-sid的列号
    sid_start_row : 读取excel-sid的起始行号
    write_col : 记录写入的列号
    write_start_row : 记录写入的行号
    homework_result=[] : 没交作业的学生学生列表r
    '''
    def set_homework_result(self, path, sid_col=0, sid_start_row=0, write_col=0, write_start_row=0, homework_result=[]):
        mapping = set(homework_result)
        r_excel_data = xlrd.open_workbook(path, formatting_info=True)  # 通过xlrd获取excel
        w_excel_data = xlutils.copy.copy(r_excel_data)  # 通过xlutils.copy的copy函数copy一份可以编辑的excel
        w_sh = w_excel_data.get_sheet(0)    # 可编辑的sheet
        r_sh = r_excel_data.sheet_by_index(0)  # 可读的sheet
        assert sid_start_row == write_start_row
        for row in range(write_start_row, r_sh.nrows):
            sid = r_sh.cell_value(row, sid_col)
            if sid not in mapping: # 如果学生不在未在没交作业的文件
                w_sh.write(row, write_col, '1')
        w_excel_data.save(path)

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
                file_list = glob.glob(os.path.join(student_dir, '**', type), recursive=True)  # 对file_list进行for循环, search/app/src/
                for file_path in file_list:
                    if (self._check_is_not_MainActivity(type, file_path, '/app/src', exclude = exclude)): \
                            # 如果不是在app/src目录下的.java文件, continue
                        continue
                    try:
                        shutil.copy(file_path, student_dir)
                    except shutil.SameFileError:  # 如果目录一一致, 则跳过
                        pass

            if remove is True:  # 如果为remove = True, 删除文件夹
                self.delete_dir(student_dir)

    '''
    解压rar文件到指定目录
    path : 解压的目的目录
    types : 解压文件类型
    '''
    def unpack(self, path, types=['*.zip', '*.rar'], deep=0, dir_sid=[]):  # 写得有点问题
        for type in types:
            assert self._dist_path != None  # 先设置好dist_path
            self._search_path = self._set_glob_search_file(path, type)
            files = sorted(self._get_specific_file_list(True))  # 递归获取列表
            if len(files) == 0:
                continue  # 没有则下一个循环
            regex_pattern = re.compile(r'\d{8}')  # 定义正则pattern
            for file in files:
                # 其实这里是否需要加入assert
                print(file + '    正在解压')
                if deep == 0:
                    sid = regex_pattern.findall(file)
                else:
                    sid = dir_sid
                if not len(sid) == 1:  # 如果获取到的sid不等于1
                    print('正则匹配出现异常, 该文件名为 {}'.format(file))
                    continue
                student_dir = os.path.join(self._dist_path, sid[0]) if deep == 0 else os.path.join(self._dist_path, sid[0], str(deep))
                # print('student_dir : ', student_dir, '   deep :', deep)
                try:
                    os.makedirs(student_dir)  # 创建多级目录
                except FileExistsError:  # 文件已经存在
                    pass
                try:
                    compressed_object = rarfile.RarFile(file) if file.endswith('.rar') else zipfile.ZipFile(file) # 可拓展性不强
                    compressed_object.extractall(student_dir)
                    # 递归继续搜索
                    self.unpack(student_dir, types, deep+1, sid)
                except:  # 避免一些错误的压缩文件
                    pass

    '''
    检查每位学生文件夹是否存在所需文件
    path : 学生集文件夹路径
    required_files : 必要的文件后缀， 格式为.xxx,(不要缺少.)
    extended_files : 可拓展的文件后缀， 为dict类型， key为扩展文件后缀， value为对应映射的文件后缀， 如{'doc': 'pdf'}， 表示可支持doc, 对应必要文件pdf
    '''
    def check_file_is_missing(self, path, required_files=[], extended_files=dict()):
        dirs = os.listdir(path)
        for dir in dirs:
            self._check_missing(os.path.join(path, dir), required_files, extended_files)

    '''
    检查每位学生文件夹情况， 暂时只打印
    file_path : 学生文件夹路径
    '''
    def _check_missing(self,  file_path, required_files, extended_files):
        files = os.listdir(file_path)
        mapping = set(required_files)
        for file in files:
            split_result = os.path.splitext(file)
            if(len(split_result) > 1):  # 避免有dir, 和一些.xxx文件异常
                if split_result[1] in mapping:
                    mapping.remove(split_result[1])
                if split_result[1] in extended_files and extended_files[split_result[1]] in mapping:  # 扩展名在extened_files上且拓展项对应的文件名也在mapping中
                    mapping.remove(extended_files[split_result[1]])
        if len(mapping) > 0:  # 仍有需要的文件没搜索到， 打印结果
            print('学号 :' + os.path.basename(file_path) + '     缺少文件', mapping)

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
        return os.path.join(path, '**', file_type)  # 方便跨平台

    def _get_specific_file_list(self, recursive):  # 获取文件列表
        return glob.glob(self._search_path, recursive=recursive)  # 是否递归

    def _check_is_not_MainActivity(self, type, file_path, specific_dir = None, exclude=[]):  # 写得有点问题
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
