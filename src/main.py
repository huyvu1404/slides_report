from shapes import *
from text_box import *
from presentations import *
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

def main():
    prs = init()
    create_first_slide(prs)
    create_second_slide(prs)
    create_third_slide(prs)
    create_fourth_slide(prs)
    create_fifth_slide(prs)
    create_sixth_slide(prs)
    prs.save('slides/slides2.pptx')

if __name__ == "__main__":
    main()