if __name__ == '__main__':
    from tkinter import *
    from os.path import expanduser
    from tkinter.messagebox import showinfo
    from tkinter.filedialog import askopenfilename, askdirectory
    det, list2, master, file, dir, home, cont, up = ([], [], Tk(), '', '.', expanduser('~') + '/', True, [None, None])
    list1 = ('Acadamic Year', 'Exam Type', 'Branch', 'Subject', 'Date', 'Exam Max Marks',
             'Exam Time duration', 'Submit Time duration')

    def get_entry_fields():
        global cont
        global det
        global dir
        global file
        det = []
        for elem in list2:
            string = elem.get()
            if string == '':
                det = []
                showinfo(message='Fill above all details')
                return
            det.append(string)

        for iii in range(1, 4):
            for jjj in det[iii]:
                if jjj in ('_', '/', '\\', ':'):
                    showinfo(message="Don't use [ _ , / , \\ , : ] characters for " + list1[iii])
                    return

        try:
            if not file.endswith('.docx'):
                raise Exception()
        except:
            showinfo(message='Select a .docx file')
            return
        else:
            if dir == '':
                showinfo(message='Fill .jar Directory field')
                return

        cont = False
        master.quit()
        with open(home + '.Old_docx_data.txt', 'w') as (fff):
            for elem in det:
                fff.write(elem + '\n')

            fff.write(file + '\n')
            fff.write(dir + '\n')


    def show_file_dialog():
        global file
        file = askopenfilename(initialdir='.', title='Select file', filetypes=(('Docx files', '*.docx'), ('all files', '*.*')))
        up[0].config(text=file)


    def show_dir_dialog():
        global dir
        dir = askdirectory(initialdir='.', title='Select Directory')
        up[1].config(text=dir)


    selected, chb = False, None

    def retrieve():
        global dir
        global file
        global selected
        if not selected:
            try:
                i, fff = 0, open(home + '.Old_docx_data.txt', 'r')
                data = fff.readlines()
                while i < 8:
                    list2[i].delete(0, END)
                    list2[i].insert(0, data[i][:-1])
                    i += 1

                selected, file, dir = True, data[8][:-1], data[9][:-1]
                up[0].config(text=file)
                up[1].config(text=dir)
                fff.close()
                return
            except:
                showinfo(message="We don't have backup data, plz type details")
                chb.deselect()

        file, dir, i, selected = (
         '', '', 0, False)
        while i < 8:
            list2[i].delete(0, END)
            i += 1

        up[0].config(text=file)
        up[1].config(text=dir)


    chb = Checkbutton(master, text='  Retrieve previous docx data', variable=IntVar(), command=retrieve)
    chb.grid(column=1, row=0, pady=10)
    for i in range(8):
        Label(master, text=list1[i]).grid(row=i + 1, padx=5, pady=6)

    for i in range(8):
        list2.append(Entry(master))
        list2[i].grid(row=i + 1, column=1, padx=5, pady=5)

    Button(master, text='Upload .docx file', command=show_file_dialog).grid(row=14, column=0, sticky=W, padx=5, pady=5)
    Button(master, text='Upload .jar Directory', command=show_dir_dialog).grid(row=15, column=0, sticky=W, padx=5, pady=5)
    Button(master, text='Quit', command=exit).grid(row=16, column=0, sticky=W, padx=5, pady=5)
    Button(master, text='Create Jar', command=get_entry_fields).grid(row=16, column=1, sticky=W, padx=5, pady=5)
    up[0], up[1] = Label(master, text=''), Label(master, text='')
    up[0].grid(row=14, column=1)
    up[1].grid(row=15, column=1)
    mainloop()
    if cont:
        quit()
    from sys import exit
    from shutil import rmtree
    from random import randint
    from zipfile import ZipFile
    from base64 import b64encode
    from os import system, listdir, remove, rename, mkdir
    charan, map, sum, Question, file1, file2, list, termi, unicode, key = (
     0, 0, 0, 0, None, None, None, chr(0), 0, None)
    dict, tuple, chrs = ('#Q', '#q', '#O', '#S', '#s', '#C', '#D'), ('\\', ';', '@'), ('Q',
                                                                                       'q',
                                                                                       'O',
                                                                                       'S',
                                                                                       's',
                                                                                       'C',
                                                                                       'D')
    zz, zzz = (0, 0)

    def trim():
        global charan
        global list
        global zz
        global zzz
        zz, loop = 0, len(list[charan])
        while zz < loop and list[charan][zz] == ' ':
            zz += 1

        zzz = zz + 2


    def File(str=None, x=zzz, bool=True, EOF=True):
        global file1
        global key
        global unicode
        if str == None:
            str = list[charan]
        map = len(str)
        while x < map:
            if x and not x % 100:
                file1.write(chr(ord(';') ^ key))
            if bool and str[x] in tuple:
                file1.write(chr(ord('\\') ^ key))
            try:
                file1.write(chr(ord(str[x]) ^ key))
            except:
                unicodechr = str[x].encode('utf-8')
                unicode += len(unicodechr) - 1
                file1.write(chr(ord('@') ^ key))
                file1.write(unicodechr)
                file1.write('@')
            else:
                if bool and str[x] == '#' and str[(x + 1)] in chrs:
                    file1.write(chr(ord('\\') ^ key))
                x += 1

        if EOF:
            file1.write(chr(ord(';') ^ key))
        return


    def iter(hasattr):
        global charan
        global file2
        global map
        global sum
        global zzz
        file2.write(str(file1.tell() - unicode ^ key) + termi)
        while True:
            File(x=zzz)
            charan += 1
            if charan == map:
                break
            trim()
            ttr = list[charan][zz:zzz]
            if ttr in dict:
                break
            if ttr == '#I':
                sum += 1
                File(str='#Iimage' + str(sum) + '.png', bool=False)
            else:
                zzz = 0

        File(str=hasattr, x=0, bool=False)


    def readContent():
        global Question
        global charan
        global file1
        global file2
        global key
        global list
        global map
        global sum
        global unicode
        try:
            import xml.etree.ElementTree as ET
            import zipfile as zf
        except ImportError:
            print('No modules')
            quit()
        else:
            WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
            PARA = WORD_NAMESPACE + 'p'
            TEXT = WORD_NAMESPACE + 't'
            list = []
            

            z = zf.ZipFile(file)
            f = z.open("word/document.xml")   # a file-like object
            tree = ET.parse(f)                # an ElementTree instance

            for elem in tree.iter():
                if elem.text:
                    list.append(elem.text)

            print(list)
            # for paragraph in tree.getiterator(PARA):
            #     texts = [ node.text for node in paragraph.getiterator(TEXT) if node.text ]
            #     if texts:
            #         list.append(('').join(texts))

        if list == []:
            showinfo(message=file + '\nis empty file')
            quit()
        else:
            key = randint(40, 60)
            file1, file2, charan, map, iii = (open('image.png', 'w'), open('image0.png', 'w'), 0, len(list), 0)
            for lop in range(8):
                file1.write(chr(randint(0, 100)))

            file1.write(chr(int(det[6])))
            file1.write(chr(int(det[7])))
            file1.write(chr(randint(0, 100)))
            file1.write(chr(key * 2))
            file1.write(chr(randint(0, 100)))
            while iii < 6:
                File(str=det[iii].strip(), x=0, bool=False)
                iii += 1

            temp = []
            while True:
                trim()
                hasattr = list[charan][zz:zzz]
                if hasattr in dict or charan + 1 == map:
                    break
                temp.append(list[charan])
                charan += 1

            file1.write(chr(len(temp)))
            for z in temp:
                File(str=z, x=0)
                File(str='#C', x=0, bool=False)

            char, sum, Question, new, unicode, qcount = (
             'S', 0, 0, True, 0, 0)
            while charan < map:
                trim()
                hasattr = list[charan][zz:zzz]
                if hasattr == '#Q' or hasattr == '#q':
                    qcount += 1
                    if new:
                        file2.write('\n')
                    else:
                        new = True
                    Question += 1
                    iter(hasattr)
                elif hasattr == '#C':
                    new = False
                    file2.write('\nC')
                    iter('#C')
                elif hasattr == '#O':
                    iter('#O')
                elif hasattr == '#I':
                    sum += 1
                    File(str='#Iimage' + str(sum) + '.png', bool=False)
                elif hasattr == '#S':
                    file2.write('\n' + termi + str(file1.tell() - unicode ^ key) + '\nC')
                    file1.write(chr(ord(char) ^ key))
                    char = 'S'
                    iter('#C')
                elif hasattr == '#s':
                    file2.write('\n' + termi + str(file1.tell() - unicode ^ key) + '\nC')
                    file1.write(chr(ord(char) ^ key))
                    char = 's'
                    iter('#C')
                elif hasattr == '#D':
                    file2.write('\n' + termi + str(file1.tell() - unicode ^ key) + '\n\x00C')
                    file1.write(chr(ord(char) ^ key))
                    char = 'D'
                    iter('#C')

            file2.write('\n' + termi + str(file1.tell() - unicode ^ key))
            file1.write(chr(ord(char) ^ key))
            file1.seek(12, 0)
            file1.write(chr(qcount))
            file1.close()
            file2.close()


    def createJar():
        try:
            mkdir('META-INF')
        except:
            rmtree('META-INF')
            mkdir('META-INF')

        with open('META-INF/MANIFEST.MF', 'w') as (file1):
            file1.write('Class-Path: .\nMain-Class: Charan\n')
        filename = dir + '/' + det[0] + '_' + det[1] + '_' + det[2] + '_' + det[3] + '.jar'
        _ALL_ = ('META-INF/MANIFEST.MF', 'image.png', 'image0.png', 'Charan.class',
                 'a.class', 'b.class', 'rgukt.png')
        with ZipFile(filename, 'w') as (zipFile):
            for elem in _ALL_:
                zipFile.write(elem)

            if sum:
                zip = ZipFile(file, 'r')
                zip.extractall('Zip')
                zip.close()
                list, charan = listdir('Zip/word/media'), 0
                while charan < sum:
                    IMAGE = 'Zip/word/media/' + list[charan]
                    with open(IMAGE, 'rb') as (imageFile):
                        str = b64encode(imageFile.read())
                    with open(list[charan], 'wb') as (imageFile):
                        imageFile.write(str)
                    zipFile.write(list[charan])
                    charan += 1

        print ('\n\n' + file + '   -    created successfully')


    readContent()
    createJar()
    try:
        rmtree('Zip')
    except OSError as e:
        pass
    else:
        try:
            rmtree('META-INF')
        except OSError as e:
            pass

    remove('image.png')
    remove('image0.png')
    IM_LIST = [ data for data in listdir('.') if data.endswith('.png') or data.endswith('.jpeg') ]
    IM_LIST.remove('rgukt.png')
    for ims in IM_LIST:
        remove(ims)