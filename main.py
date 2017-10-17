from unpack_Tools import *
import numpy as np

if __name__ == '__main__':
    tool = Tools()
    tool.init_student_list('/home/ysing/Documents/TA-MAD/名单/17.1手机平台（理论+实验）名单.xls')
    result = tool.get_homework_result(['/home/ysing/Documents/TA-MAD/Homework-zip/Lab-4-zip/**/*.zip', \
                                       '/home/ysing/Documents/TA-MAD/Homework-zip/Lab-4-zip/**/*.rar'])

    for row in range(0, len(result), 3):
        print(result[row:row+3])

    # tool.create_dist_dir('/home/ysing/Documents/TA-MAD/Homework-zip/Lab-4-Homework')
    # tool.unpack_zip('/home/ysing/Documents/TA-MAD/Homework-zip/Lab-4-zip/**/*.zip') #搜索zip路径
    # tool.copy_specific_type('/home/ysing/Documents/TA-MAD/Homework-zip/Lab-4-Homework', \
    #                         remove=True, types = ['*.pdf', '*.java'], exclude=['ExampleUnitTest.java', 'ExampleInstrumentedTest.java'])