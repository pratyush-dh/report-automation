import zipfile
import shutil
import os
from PIL import Image

def unzip_file(template_report_file, destination):
    """
    Unzips a docx file to the specified destination folder.

    Args:
        template_report_file (str): Path to the input docx file.
        destination (str): Destination folder where the files will be extracted.

    Returns:
        None
    """
    parent_dir = os.path.dirname(os.path.abspath(template_report_file))
    zip_obj = zipfile.ZipFile(template_report_file, 'r')
    zip_obj.extractall(os.path.join(parent_dir, destination))
    zip_obj.close()
    print('Unzip successful!')

def replace_images(new_images_folder, images_folder):
    """
    Replaces the images in the docx file with new images.

    Args:
        new_images_folder (str): Folder containing new images.
        images_folder (str): Folder containing the original images.

    Returns:
        None
    """
    print('Starting image replacement.')
    for image in os.listdir(new_images_folder):
        shutil.copy(os.path.join(new_images_folder, image), os.path.join(images_folder, image))
        print(f'\tReplaced image {image}.')
    print('Image replacement successful.')

def recreate_docx(destination, folder_name):
    """
    Converts the folder with updated files to a docx file.

    Args:
        destination (str): Folder containing the updated files.
        folder_name (str): Name of the resulting docx file.

    Returns:
        None
    """
    shutil.make_archive(folder_name, 'zip', destination)
    os.rename(folder_name + '.zip', folder_name + '.docx')
    print(f"Zip file '{folder_name}.docx' created successfully!")
    shutil.rmtree(destination)

def rename_images(original_folder, rename_dict):
    """
    Renames images in the original folder according to the provided dictionary.

    Args:
        original_folder (str): Folder containing the original images.
        rename_dict (dict): Dictionary mapping old image names to new names.

    Returns:
        None
    """
    renamed_images_folder = '_converted'
    if not os.path.exists(renamed_images_folder):
        os.mkdir(renamed_images_folder)

    for old_name, new_name in rename_dict.items():
        new_path = os.path.join(renamed_images_folder, new_name)
        image = Image.open(os.path.join(original_folder, old_name))
        image.save(new_path)
        image.close()

def fully_update_docx(destination, template_report_file, new_images_folder, report_folder, rename_dict):
    """
    Fully updates a docx file by unzipping, renaming images, replacing images, and recreating the docx.

    Args:
        destination (str): Folder where the docx contents are extracted and updated.
        template_report_file (str): Path to the original docx file.
        new_images_folder (str): Folder containing new images.
        report_folder (str): Folder where the final docx will be saved.
        rename_dict (dict): Dictionary mapping old image names to new names.

    Returns:
        None
    """
    renamed_images_folder = '_converted'
    parent_dir = os.path.dirname(os.path.abspath(template_report_file))
    images_folder = os.path.join(parent_dir, destination, 'word', 'media')
    
    print(f'images folder: ', images_folder)
    
    folder_name = os.path.basename(destination)

    unzip_file(template_report_file, destination)
    rename_images(new_images_folder, rename_dict)
    replace_images(renamed_images_folder, images_folder)
    recreate_docx(destination, folder_name)
    print('File update successful.')

    if not os.path.exists(report_folder):
        os.mkdir(report_folder)

    final_report_path = os.path.join(report_folder, folder_name + '.docx')
    shutil.move(folder_name + '.docx', final_report_path)
    print(f'Report saved to "{final_report_path}".')
    
def create_name_dict(name_txt_file):
    """
    Creates a dictionary from a text file containing image name mappings.

    Args:
        name_txt_file (str): Path to the text file.

    Returns:
        dict: Dictionary mapping old image names to new names.
    """
    with open(name_txt_file, "r") as f:
        image_dict = {}
        for line in f:
            line = line.strip().split(",")
            if len(line) == 2:
                key = line[0].strip()
                value = line[1].strip()
                image_dict[key] = value
            else:
                print(f"Invalid line: {line}")
        return image_dict
