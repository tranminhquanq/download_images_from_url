import requests
import os
import xlsxwriter
import xlrd
# pip3 install xlrd==1.2.0

image_urls = []
broken_images = []
valid_image = [".jpg", ".png", "jpeg"]

# Give the location of the file
loc = ("data.xlsx")
# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)

for i in range(sheet.nrows):
    image_urls.append(sheet.cell_value(i, 0))


workbook = xlsxwriter.Workbook('images.xlsx')
worksheet = workbook.add_worksheet()

# format style
worksheet.set_column('A:A', 30)
worksheet.set_column('B:B', 30)
worksheet.set_default_row(90)
bold = workbook.add_format({'bold': True})
worksheet.write('A1', "Tên", bold)
worksheet.write('B1', "Ảnh main", bold)
worksheet.write('C1', "Ảnh 1", bold)
worksheet.write('D1', "Ảnh 2", bold)


def save_image(image_path, image_name, position):
    worksheet.write('A' + str(position + 1), image_name, bold)
    worksheet.insert_image('B' + str(position + 1), image_path,
                           {'x_scale': 0.15, 'y_scale': 0.15})
    # if image_index == 0:
    #     worksheet.insert_image('B' + str(position + 1), image_path,
    #                            {'x_scale': 0.15, 'y_scale': 0.15})
    # elif image_index == 1:
    #     worksheet.insert_image('C' + str(position + 1), image_path,
    #                            {'x_scale': 0.15, 'y_scale': 0.15})
    # elif image_index == 2:
    #     worksheet.insert_image('D' + str(position + 1), image_path,
    #                            {'x_scale': 0.15, 'y_scale': 0.15})


def download(url, file_name, name, old_path, new_path):
    r = requests.get(url, stream=True)
    if r.status_code == 200:
        # This command below will allow us to write the data to a file as binary:
        with open(file_name, 'wb') as f:
            for chunk in r:
                f.write(chunk)
            os.rename(old_path, new_path)
            save_image(new_path, name, i)
    else:
        # We will write all of the images back to the broken_images list:
        broken_images.append(url)


def download_images(urls):
    valid_image = [".jpg", ".png", "jpeg"]
    for i in range(1, len(urls)):
        url = urls[i].replace("wid=3000&hei=3000", "wid=600&hei=600", 1)
        # We can split the file based upon / and extract the last split within the python list below:
        file_name = url.split("/")[-1]
        name = url.split("/")[-1].split("_")[0]
        print(file_name)
        if "_main" in file_name:
            print(file_name)

            if file_name[-4:] in valid_image:
                # Now let's send a request to the image URL:
                old_path = "./" + file_name
                new_path = "./all_images/" + file_name
                download(url, file_name, name, old_path, new_path)
                url = url.replace("_main", "_1", 1)

            else:
                file_name = file_name + ".jpeg"
                old_path = "./" + file_name
                name = url.split("/")[-1].split("_")[0]
                new_path = "./all_images/" + file_name
                download(url, file_name, name, old_path, new_path)
                url = url.replace("_main", "_1", 1)

    # workbook.close()


# image_urls = [
#     '',
#     'https://fossil.scene7.com/is/image/FossilPartners?wid=3000&hei=3000&layer=0&src=is%7BFossilPartners/SKW2972_main%7D']


download_images(image_urls)
