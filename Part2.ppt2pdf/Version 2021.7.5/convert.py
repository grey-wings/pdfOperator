import os
import comtypes.client


# 文档地址：https://pythonhosted.org/comtypes/
def init_powerpoint():
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    # print(help(powerpoint))
    return powerpoint


def ppt_to_pdf(powerpoint, inputFileName, outputFileName, formatType = 32):
    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName + ".pdf"
    deck = powerpoint.Presentations.Open(inputFileName)
    # comtypes.POINTER（Presentations）实例的Open（…）方法
    deck.SaveAs(outputFileName, formatType)  # formatType = 32 for ppt to pdf
    deck.Close()


def convert_files_in_folder(powerpoint, folder):
    files = os.listdir(folder)
    pptfiles = [f for f in files if f.endswith((".ppt", ".pptx"))]
    for pptfile in pptfiles:
        fullpath = os.path.join(folder, pptfile)
        ppt_to_pdf(powerpoint, fullpath, fullpath)


def batch_convert(powerpoint, fileList, saveList):
    if not saveList:
        saveList = [i.partition('.')[0] for i in fileList]
    for i in range(len(fileList)):
        ppt_to_pdf(powerpoint, fileList[i], saveList[i])


if __name__ == "__main__":
    powerpoint = init_powerpoint()
    cwd = os.getcwd()
    convert_files_in_folder(powerpoint, cwd)
    powerpoint.Quit()
