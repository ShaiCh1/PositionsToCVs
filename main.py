import docx2txt
from zipfile import ZipFile
import pickle
import os
import pandas as pd
import matplotlib.pyplot as plt

### Get files from zip
def ReadZipFile(file_name):
    with ZipFile(file_name, 'r') as zip:
        #zip.printdir()  #printing all the contents of the zip file
        print('Extracting all the files now...')
        zip.extractall() # extracting all the files
        print('Done!')

        with ZipFile(file_name, 'r') as zipObj:
            listOfiles = zipObj.namelist() # Get list of files names in zip

            totalFiles = 0
            for elem in listOfiles: #count number of files in zip
                totalFiles += 1
        return totalFiles

###import CVs (Docx file from library)
def Import_CVs_Files(CV_index):
    path = "CVs"
    CVs_list = os.listdir(path)
    #print("Files and directories in '", path, "' :")
    #print(CVs_list)
    CV_Name = CVs_list[CV_index]
    filename = "CVs\\" + CV_Name
    CV_contents = docx2txt.process(filename) # Store the CV in a variable
    #print(CV)
    return CV_Name,CV_contents

###import positions (pkl files)
def Import_Position_Files(Item_index):
    with open(Item_index, 'rb') as item:
      job_description = pickle.load(item) # Store the job_description in a variable
    #print("job des:",job_description)  #Print the job_description
    Position = str(job_description)
    return Position

###Calculate file match percentages
def CalcMatch(CV,Position):
    text = [CV, Position] #list of text CV + job description

    ##Perform the test to find similarities between the documents
    from sklearn.feature_extraction.text import CountVectorizer
    cv = CountVectorizer()
    count_matrix = cv.fit_transform(text) #Create a count vectorizer object and convert the text documents to a matrix of token counts.

    from sklearn.metrics.pairwise import cosine_similarity


    #print("\nSimilarity Scores:") #Print the similarity scores
    #print(cosine_similarity(count_matrix)) #Get and print the cosine similarity scores.

    matchPercentage = cosine_similarity(count_matrix)[0][1] * 100
    matchPercentage = round(matchPercentage, 2) #round to two decimal
    #print("Your CV matches about " + str(matchPercentage)+ "% of the Position.") #get the match percentage
    return matchPercentage

def main():
    TotalNumCV = ReadZipFile("CVs.zip")  # get number of CV files and extract zip
    TotalNumPos = ReadZipFile("positions.zip") #get number of position files and extract zip
    df = pd.DataFrame(columns= ['CV'])

    for i in range(TotalNumCV):
        print(i)
        tmpInd = 1
        CV_Name,CV_contents = Import_CVs_Files(i)
        df.at[i,'CV'] = CV_Name
        for j in range(TotalNumPos):
            Position_Name = "item_"+str(j)
            Position = Import_Position_Files(Position_Name+".pkl")
            matchPercentage = CalcMatch(CV_contents,Position)
            Position_Header = "Position_Name_" + str(tmpInd)
            if matchPercentage > 55:
                if Position_Header not in df.columns:
                    df['Position_Header'] = ''
                df.at[i, Position_Header] = Position_Name
                tmpInd += 1



    df.head()
    df.to_csv('Result.csv')

if __name__ == "__main__":
    main()