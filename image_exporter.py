import os
import sys
import docx2txt

def export_image(file_path):

    file_name = os.path.basename(file_path)
    file_dir = os.path.dirname(file_path)

    if file_name[len(file_name)-5:] != '.docx':
        print('just docx files')
        return
    
    image_path = file_dir + '/' + file_name[:len(file_name)-5] + '/'
    make_file(image_path)
    
    # extract text and write images in /tmp/img_dir
    text = docx2txt.process(file_path, image_path) 

def make_file(path):
    # image_path not exist this make one
    if not os.path.exists(path):
        try:
            os.makedirs(path)
        except OSError:
            print("Unable to create img_dir {}".format(path))
            sys.exit(1)


def detect(path):

    
    for item in os.listdir(path):
        item_path = os.path.join(path, item)

        if os.path.isfile(item_path):
            export_image(item_path)
        
        elif os.path.isdir(item_path):
            detect(item_path)



arg_size = len(sys.argv)

if arg_size == 1:
    print('set word file address')

elif arg_size > 1:
    if os.path.isfile(sys.argv[1]):
        export_image(sys.argv[1])
    
    elif os.path.isdir(sys.argv[1]):
        detect(sys.argv[1])





# os.system('pause')


