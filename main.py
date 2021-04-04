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
worksheet.write('B1', "Ảnh", bold)


def save_image(image_path, image_name, position):
    print(position)

    worksheet.write('A' + str(position + 1), image_name, bold)
    worksheet.insert_image('B' + str(position + 1), image_path,
                           {'x_scale': 0.15, 'y_scale': 0.15})


def download_images(urls):
    for i in range(1, len(urls)):
        url = urls[i].replace("wid=3000&hei=3000", "wid=600&hei=600", 1)
        # We can split the file based upon / and extract the last split within the python list below:
        file_name = url.split("/")[-1]
        name = url.split("/")[-1].split("_")[0]
        if file_name[-4:] in valid_image:
            # print(f"This is the file name: {file_name}")
            # Now let's send a request to the image URL:
            old_path = "./" + file_name
            new_path = "./all_images/" + file_name

            r = requests.get(url, stream=True)
            # We can check that the status code is 200 before doing anything else:
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

        else:
            file_name = file_name + ".jpeg"
            old_path = "./" + file_name
            name = url.split("/")[-1].split("_")[0]
            new_path = "./all_images/" + file_name
            r = requests.get(url, stream=True)
            if r.status_code == 200:
                with open(file_name, 'wb') as f:
                    for chunk in r:
                        f.write(chunk)

                os.rename(old_path, new_path)
                save_image(new_path, name, i)

            else:
                broken_images.append(url)
    workbook.close()


# image_urls = ['https://sempioneer.com/wp-content/uploads/2020/05/dataframe-300x84.png',
#               'https://sempioneer.com/wp-content/uploads/2020/05/json_format_data-300x72.png',
#               'https://c.files.bbci.co.uk/12A9B/production/_111434467_gettyimages-1143489763.jpg',
#               'https://i.ytimg.com/vi/jHWKtQHXVJg/maxresdefault.jpg',
#               'https://img.webmd.com/dtmcms/live/webmd/consumer_assets/site_images/article_thumbnails/other/cat_relaxing_on_patio_other/1800x1200_cat_relaxing_on_patio_other.jpg',
#               'https://i.natgeofe.com/n/9135ca87-0115-4a22-8caf-d1bdef97a814/75552.jpg?w=636&h=424']


download_images(image_urls)
