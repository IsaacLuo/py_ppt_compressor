import sys
import zipfile
import conf
import os
import datetime
import re
import subprocess
import shutil

if len(sys.argv) > 1:
    ppt_file_name = sys.argv[1]
else:
    print('usage: python3 main.py your_slides.pptx')

tmp_dir = os.path.abspath(conf.TMP_DIR)
tmp_dir = os.path.join(tmp_dir, ppt_file_name+'_'+datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S.%f'))
os.mkdir(tmp_dir)

with zipfile.ZipFile(ppt_file_name, 'r') as zip_ref:
    zip_ref.extractall(tmp_dir)

media_folder = os.path.join(tmp_dir,'ppt','media')
media_files = os.listdir(media_folder)

target_media_file_name_re = re.compile(r'\.(emf|tiff)$')

media_files = [x for x in media_files if target_media_file_name_re.findall(x) != []]

converted_file_names = []

for media_file_name in media_files:
    print(media_file_name)
    fn1, fn2 = os.path.splitext(media_file_name)
    fn2_dst = '.jpg' if fn2 == '.emf' else '.png'
    convert_result = subprocess.call('"{}" "{}" "{}"'.format(
        os.path.join(conf.IMAGE_MAGIC_DIR, 'convert'),
        os.path.join(media_folder, media_file_name),
        os.path.join(media_folder, fn1+fn2_dst)))
    if convert_result == 0:
        os.remove(os.path.join(media_folder, media_file_name))
        converted_file_names.append((media_file_name, fn1+fn2_dst, re.compile('media/'+media_file_name), 'media/'+fn1+fn2_dst))

# search xml to change all pictures path
res_folder = os.path.join(tmp_dir,'ppt','slides', '_rels')

for file_name in os.listdir(res_folder):
    text = ''
    print(file_name)
    with open(os.path.join(res_folder, file_name), 'r+', encoding='utf-8') as f:
        text = f.read()
        for item in converted_file_names:
            text = item[2].sub(item[3], text)
        f.seek(0)
        f.truncate()
        f.write(text)


def dfs_get_zip_file(input_path,result):
    files = os.listdir(input_path)
    for file in files:
        if os.path.isdir(input_path+'/'+file):
            dfs_get_zip_file(input_path+'/'+file,result)
        else:
            result.append(input_path+'/'+file)

def zip_path(input_path,output_path,output_name):
    f = zipfile.ZipFile(output_path+'/'+output_name,'w',zipfile.ZIP_DEFLATED)
    filelists = []
    dfs_get_zip_file(input_path,filelists)
    for file in filelists:
        f.write(file)
    f.close()
    return output_path+r"/"+output_name

ppt_path, ppt_name = os.path.split(ppt_file_name)

ori_dir = os.path.abspath(os.curdir)
os.chdir(tmp_dir)
zip_path('.', ori_dir if ppt_path=='' else ppt_path, os.path.splitext(ppt_name)[0]+'_compressed.pptx')
os.chdir(ori_dir)
shutil.rmtree(tmp_dir)