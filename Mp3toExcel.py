import mutagen,xlrd, glob,re,openpyxl,os,pygal
#from mutagen.easyid3 import EasyID3
from os import walk
from pprint import pprint 
from tinytag import TinyTag, TinyTagException
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from mp3_tagger import MP3File
from string import ascii_uppercase
from mutagen.mp3 import MP3
list = os.listdir('C:\\Users\\mrdrj\\Desktop\\sdf\\') # directory path of files
number_files = len(list) +1
#print (number_files) 
#vowels =['a','e','i', 'o','u']
#index = vowels.index('e')
#print('The index of e:', index)

print(os.getcwd()) 
if os.path.isfile('C:\\Users\\mrdrj\\Desktop\\cqq\\Brimstone CD 3\\'):
                        print ("File exist")
else:
                        print ("File not exist")
                        
def ExtractMP3TagtoExcel():


    
    
    print('Starting Program')
    
    tracks= []
    gettags =[]
    getit = []

    for root, dirs, files, in os.walk ('C:\\Users\\mrdrj\\Desktop\\sdf\\'):
        for name in files:
            if name.endswith(('.mp3','.m4a','.flac','.alac')):

                tracks.append(name) #Add Media Files

                try:

                    temp_track = TinyTag.get(root + '\\' + name)
                    mp3 = MP3File(root + '\\' + name)
                    #tags = mp3.get_tags()
                    print(root, '-',temp_track.artist, '-', temp_track.title)

                    gettags2 = [temp_track.album, temp_track.albumartist, temp_track.artist, temp_track.audio_offset,
                                temp_track.bitrate, temp_track.comment, temp_track.composer, temp_track.disc,
                                temp_track.disc_total, temp_track.duration, temp_track.filesize, temp_track.genre,
                                temp_track.samplerate, temp_track.title, temp_track.track, temp_track.track_total,
                                temp_track.year] #Add Tags to list
                    print('----' * 20)

                  
    
                    for x in range(len(gettags2)):
                    #append slice of gettags2, containing the entire gettags2
                        gettags.append(gettags2[:])
                        #print(gettags2[x]) 
                except TinyTagException:
                    print('Error')
                

   
    wb = Workbook()
    os.chdir('C:\\Users\\mrdrj\\Desktop\\cqq\\Brimstone CD 3\\')
   
    dest_filename =   'empty_book.xlsx'
    newFile = dest_filename
    worksheet = wb.active
    wb = openpyxl.load_workbook(filename = newFile)  

    ws1 = wb.active
    ws = wb.active
    ws1.title = "MP3 Info" # Main Tab
    sheet = "MP3 Info"

    #print(Mp3Tagsexel)
    #Add Columns to Document
    ws1['A1'] = 'Album'
    ws1['B1'] = 'Contributing Artists'
    ws1['C1'] = 'Title'
    ws1['E1'] = 'Genre'
    ws1['F1'] = 'Disc Number'
    ws1['G1'] = 'Track Duration'
    for col in range(1, 2): # Add how many Tabs 
     #ws1.append(range(5)) #Add values to Rows

        for row in range(1, number_files): 
            #for col in range(1, 8): # Number of colums static
                #print('------View Results------')
                      
           #for i in range(1, 2):
              #print(i, ws1.cell(row=i, column=2).value)
            #for i in range(1, 11):         # create some data in column A
            sheet['A' + str(i)] = i

            for r in gettags:
                ws1['A' + len(str(r))] = r[0]
                #_ = ws1.cell(column=col, row=row, value= gettags2[1]) #"{0}".format(get_column_letter(col)))
                      
                    ##column_cell = 'A'
                    #ws1['A1'] = 'Album'
                    ws1[column_cell + str(row + 1)] = r[0]

                    column_cell = 'B'
                    #ws1['B1'] = 'Contributing Artists'
                    ws1[column_cell + str(row + 1)] =  r[1]    

                    column_cell = 'C'
                    #ws1['C1'] = 'Title'
                    ws1[column_cell + str(row + 1)] = r[2] 

                    column_cell = 'D'
                    #ws1['D1'] = 'Total Number of Disk'
                    ws1[column_cell + str(row + 1)] = r[3] 

                    column_cell = 'E'
                    #ws1['E1'] = 'Genre'
                    ws1[column_cell + str(row + 1)] =  r[4] 

                    column_cell = 'F'     
                    #ws1['F1'] = 'Disc Number'
                    ws1[column_cell + str(row + 1)] =  r[5] 

                    column_cell = 'G'
                    #ws1['G1'] = 'Track Duration'
                    ws1[column_cell + str(row + 1)] =  r[6] 

                    #print(r[6])
                    
                    
    wb.save(filename=dest_filename)
                

ExtractMP3TagtoExcel()
