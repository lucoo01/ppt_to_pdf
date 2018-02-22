import comtypes.client
import os

def init_powerpoint():
    powerpoint = comtype.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    return powerpoint

def ppt_to_pdf(powerpoint, inputFileName, outputFileName, formatType = 32):
    if outputFileName[-3:] != 'pdf':
        outputFileName += ".pdf"
    phander = powerpoint.Presentations.Open(inputFileName)
    phander.SaveAs(outputFileName, formatType) # 类型 32 是将 ppt 转成 pdf 
    phander.Close()

def cover_files_in_folder(powerpoint, folder):
    files = os.listdir()
    pptfiles = [f for f in files if f.endswidth(".ppt", ".pptx")]
    
    for pptfile in pptfiles:
        fullpath = os.path.join(cwd, pptfile)
        ppt_to_pdf(powerpoint, fullpath, fullpath)
    

if __name__ == "__main__":
    powerpoint = init_powerpoint()
    cwd = os.getcwd()
    cover_files_in_folder(powerpoint, cwd)
    powerpoint.Quit()
