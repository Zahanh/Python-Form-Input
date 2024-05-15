from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont
#imports Pil module
import PIL,os
 
# # creating image object which is of specific color
# im = PIL.Image.new(mode = "RGB", size = (192, 72),
#                            color = (255, 255, 255))
# myFont = ImageFont.load_default()
# I1 = ImageDraw.Draw(im)
# I1.text((10, 10), 'Nice Car',font = myFont, fill = (255,0,0))
# # this will show image in any image viewer
# im.show()

class label:
    """
    ASSUMING A 2" X 0.75" LABEL SIZE FOR A LABEL MAKER
    """
    def __init__(self,img_path:os.path) -> None:
        self.img_path = img_path
        img_size:tuple  = (192,72)
        color:tuple     = (255,255,255)

        self.im = PIL.Image.new(mode = "RGB", 
                                size = img_size,
                                color = color)
        self.font = ImageFont.load_default()


    def create_label(self,job_num,subdivision,lot,customer):
        """"
        JOB_NUM
        SUBDIVISION
        LOT & BLK
        CUSTOMER
        """
        fill_color:tuple = (0,0,0)

        I1 = ImageDraw.Draw(self.im)
        I1.text((10,10), text=job_num,font=self.font,fill=fill_color)
        I2 = ImageDraw.Draw(self.im)
        I2.text((75,10), text=subdivision,font=self.font,fill=fill_color)
        I3 = ImageDraw.Draw(self.im)
        I3.text((75,25), text=lot,font=self.font,fill=fill_color)
        I4 = ImageDraw.Draw(self.im)
        I4.text((75,40), text=customer,font=self.font,fill=fill_color)
        #self.im.show()
        self.save(job_num=job_num)

    def save(self,job_num:str):
        os.chdir(self.img_path)
        self.im.save(f'{job_num}_label.png')










