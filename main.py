from unpack_Tools import *

if __name__ == '__main__':
    tool = Tools()
    #tool.init_student_list('/home/ysing/Documents/TA-MAD/名单/17.1手机平台（理论+实验）名单.xls', 0, 2)
    result = tool.get_homework_result('/home/ysing/Documents/TA-MAD/Homework-zip/Lab-4-Homework')
    tool.set_homework_result('/home/ysing/Documents/TA-MAD/名单/17.1手机平台（理论+实验）名单-提交记录.xls', 0, 1, 7, 1, result)

    #for row in range(0, len(result), 3):
    #    print(result[row:row+3])

    #tool.create_dist_dir('/home/ysing/Documents/TA-MAD/Homework-zip/Lab-4-Homework')
    #tool.unpack('/home/ysing/Downloads/补交')

    #tool.copy_specific_type('/home/ysing/Documents/TA-MAD/Homework-zip/Lab-4-Homework',
    #                            remove=True, types = ['*.pdf', '*.java'], exclude=['ExampleUnitTest.java', 'ExampleInstrumentedTest.java'])
    #tool.check_file_is_missing('/home/ysing/Documents/TA-MAD/Homework-zip/Lab-4-Homework', ['.java', '.pdf'], {'.doc':'.pdf', '.docx':'.pdf'})