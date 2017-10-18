from unpack_Tools import *

if __name__ == '__main__':
    tool = Tools()
    tool.init_student_list('/home/ysing/下载/手机平台应用开发.xls', 1, 6)
    result = tool.get_homework_result('/home/ysing/PycharmProjects/TA-tool/Lab_4')
    for row in range(0, len(result), 3):
        print(result[row:row+3])

    tool.create_dist_dir('/home/ysing/Documents/TA-MAD/Homework-zip/Lab-4-Homework')
    tool.unpack('/home/ysing/PycharmProjects/TA-tool/Lab_4')
    tool.copy_specific_type('/home/ysing/Documents/TA-MAD/Homework-zip/Lab-4-Homework', \
                             remove=True, types = ['*.pdf', '*.java'], exclude=['ExampleUnitTest.java', 'ExampleInstrumentedTest.java'])