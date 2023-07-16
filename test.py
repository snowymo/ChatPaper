import numpy as np

from chat_paper_ppt import write_pptx, write_image_pptx

result_dict = {'标题': 'DigitSpace: Designing Thumb-to-Fingers Touch Interfaces (手指空间：设计拇指到手指的触摸界面)'}
images = ['export/DigitSpace Designing Thumb-to-Fingers Touch Interfaces (手指空间_设计拇指到手指的触摸界面)\\image1_2.png',  \
         'export/DigitSpace Designing Thumb-to-Fingers Touch Interfaces (手指空间_设计拇指到手指的触摸界面)\\image4_3.png',  \
         'export/DigitSpace Designing Thumb-to-Fingers Touch Interfaces (手指空间_设计拇指到手指的触摸界面)\\image4_4.png',  \
         'export/DigitSpace Designing Thumb-to-Fingers Touch Interfaces (手指空间_设计拇指到手指的触摸界面)\\image4_6.png',  \
         'export/DigitSpace Designing Thumb-to-Fingers Touch Interfaces (手指空间_设计拇指到手指的触摸界面)\\image4_7.png',  \
         'export/DigitSpace Designing Thumb-to-Fingers Touch Interfaces (手指空间_设计拇指到手指的触摸界面)\\image4_8.png',  \
         'export/DigitSpace Designing Thumb-to-Fingers Touch Interfaces (手指空间_设计拇指到手指的触摸界面)\\image5_1.png',  \
         'export/DigitSpace Designing Thumb-to-Fingers Touch Interfaces (手指空间_设计拇指到手指的触摸界面)\\image5_2.png',  \
         'export/DigitSpace Designing Thumb-to-Fingers Touch Interfaces (手指空间_设计拇指到手指的触摸界面)\\image5_3.png',  \
         'export/DigitSpace Designing Thumb-to-Fingers Touch Interfaces (手指空间_设计拇指到手指的触摸界面)\\image5_4.png',  \
         'export/DigitSpace Designing Thumb-to-Fingers Touch Interfaces (手指空间_设计拇指到手指的触摸界面)\\image5_7.png',  \
         'export/DigitSpace Designing Thumb-to-Fingers Touch Interfaces (手指空间_设计拇指到手指的触摸界面)\\image5_8.png',  \
         'export/DigitSpace Designing Thumb-to-Fingers Touch Interfaces (手指空间_设计拇指到手指的触摸界面)\\image7_1.png',  \
         'export/DigitSpace Designing Thumb-to-Fingers Touch Interfaces (手指空间_设计拇指到手指的触摸界面)\\image7_2.png',  \
         'export/DigitSpace Designing Thumb-to-Fingers Touch Interfaces (手指空间_设计拇指到手指的触摸界面)\\image7_3.png',  \
         'export/DigitSpace Designing Thumb-to-Fingers Touch Interfaces (手指空间_设计拇指到手指的触摸界面)\\image7_4.png',  \
         'export/DigitSpace Designing Thumb-to-Fingers Touch Interfaces (手指空间_设计拇指到手指的触摸界面)\\image9_1.png',  \
         'export/DigitSpace Designing Thumb-to-Fingers Touch Interfaces (手指空间_设计拇指到手指的触摸界面)\\image9_2.png']
print("images", images)
np.save("export/images.npy", images)

write_image_pptx("CHI2023_chatpaper", result_dict, images)

# print("summary time:", match_title, time.time() - start_time)