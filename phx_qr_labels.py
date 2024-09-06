import textwrap
import qrcode
from PIL import Image, ImageDraw, ImageFont
import pandas as pd


def my_textwrap(txt):
    return textwrap.wrap(txt, width=70, initial_indent='', subsequent_indent='',
                         expand_tabs=True, replace_whitespace=True, fix_sentence_endings=False,
                         break_long_words=True, drop_whitespace=True, break_on_hyphens=True,
                         tabsize=8, max_lines=None)


formulary = pd.read_excel('form_list.xlsx')
# formulary = formulary.head(20)  # for testing, remove for prod
formulary = formulary.reset_index()
for index, row in formulary.iterrows():
    data = (f"{row['PHX']};{row['Name']};{row['Dosage']}")
    if 'HIGH ALERT' in row["Flag"]:
        qr_color = 'red'
    elif 'LASA' in row["Flag"]:
        qr_color = 'orange'
    elif 'LIGHT SENSITIVE' in row["Flag"]:
        qr_color = 'green'
    else:
        qr_color = 'black'
    phx_qr = qrcode.QRCode(
        version=2,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=8,
        border=1,
    )
    sep = ';'
    phx_qr.add_data(data)
    qr_img = phx_qr.make_image(fill_color=qr_color, back_color='white')
    qr_img = qr_img.resize((290, 250))

    label_w = 1080
    label_h = 320
    my_font = "C:\Windows\Fonts\Calibri\calibri.ttf"
    my_Bold_font = "C:\Windows\Fonts\Calibri\calibrib.ttf"
    text_x_start = 240
    start = 60
    factor_low = 1.2
    factor_high = 1.8

    phx_img_1 = Image.new('RGB', (label_w, label_h), color=(255, 255, 255))
    font_phx = ImageFont.truetype(my_font, 32)
    if len(str(row['Name'])) <= 20:
        font_name = ImageFont.truetype(my_Bold_font, 70)
    elif len(str(row['Name'])) > 20 and len(str(row['Name'])) < 40:
        font_name = ImageFont.truetype(my_Bold_font, 55)
    else:
        font_name = ImageFont.truetype(my_font, 48)
    font_dosage = ImageFont.truetype(my_font, 38)
    font_flag = ImageFont.truetype(my_Bold_font, 38)

    d = ImageDraw.Draw(phx_img_1)
    shape = [(10, 10), (label_w - 10, label_h - 10)]
    d.rectangle(shape, fill=None, outline='black')
    d.text((text_x_start, 24), f"{row['PHX']}", fill=(0, 0, 0), font=font_phx)

    for line_n in textwrap.wrap(f"{row['Name']}", width=38, initial_indent='', subsequent_indent='',
                                expand_tabs=True, replace_whitespace=True, fix_sentence_endings=False,
                                break_long_words=True, drop_whitespace=True, break_on_hyphens=True,
                                tabsize=8, max_lines=3):
        d.text((text_x_start, start), line_n, fill=(0, 0, 0), font=font_name)
        start += 38*factor_high
    for line_d in textwrap.wrap(f"{row['Dosage']}", width=48, initial_indent='', subsequent_indent='',
                                expand_tabs=True, replace_whitespace=True, fix_sentence_endings=False,
                                break_long_words=True, drop_whitespace=True, break_on_hyphens=True,
                                tabsize=8, max_lines=3):
        d.text((text_x_start, start), line_d, fill=(0, 0, 0), font=font_dosage)
        start += 44*factor_low
    len_word = 30
    for word in str(row['Flag']).split(','):
        if word.strip() == 'HIGH ALERT':
            fill_f = 'red'
        elif word.strip() == 'LASA':
            fill_f = 'orange'
        elif word.strip() == 'LIGHT SENSITIVE':
            fill_f = 'green'
        else:
            fill_f = 'black'
        d.text((len_word, label_h-44),
               f'{word.strip()} ', fill=fill_f, font=font_flag)
        len_word += (len(word)*22)

    phx_img_1.paste(qr_img, (20, 20))
    phx_img_1.save(f".\codes\{row['PHX']}.png")
