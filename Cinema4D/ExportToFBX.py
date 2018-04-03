
import os

import c4d
from c4d import bitmaps, documents, gui, storage, threading


def SelectDirectory(default_path=''):
    return storage.LoadDialog(type=c4d.FILESELECTTYPE_ANYTHING ,title='Select directory with file_list.txt or to search c4d files. After then you can convert it to FBX', flags=c4d.FILESELECT_DIRECTORY, def_path=default_path).decode("utf-8")


def Export2FBX(source_file, do_render):
    loaded = documents.LoadFile(source_file)
    if not loaded:
        return (False,'LOAD: [Fail] EXPORT: [Fail] RENDER: [Fail]')
    doc = documents.GetActiveDocument()
    doc_name_blocks = doc.GetDocumentName().split('.')
    doc_name = ''
    i = 0
    while i < len(doc_name_blocks)-1:
        if len(doc_name) > 0:
            doc_name += '.'
        doc_name = doc_name + doc_name_blocks[i]
        i += 1
    dest_file = os.path.join(doc.GetDocumentPath(), doc_name)
    render_status = False
    export_status = False
    if do_render:
        renderThread = RenderFile(doc)
        renderThread.dest_file = '' + dest_file
        renderThread.Start()
        renderThread.End()
        render_status = renderThread.GetStatus()
        if render_status:
            render_status = 'Succ'
        else:
            render_status = 'Fail'
    else:
        render_status = 'Skip'
    exportThread = ExportFile(doc)
    exportThread.dest_file = '' + dest_file
    exportThread.Start()
    exportThread.End()
    export_status = exportThread.GetStatus()
    if export_status:
        export_status = 'Succ'
    else:
        export_status = 'Fail'
    return (True,'LOAD: [Succ] EXPORT: [{0}] RENDER: [{1}]'.format(export_status, render_status))

def ListC4DFiles(root_directory):
    res = []
    for root, dirs, files in os.walk(root_directory):
        for name in files:
            if os.path.splitext(name)[1][1:].strip().lower() == 'c4d':
                res.append(os.path.join(root, name))
        for name in dirs:
            full_res = ListC4DFiles(os.path.join(root, name))
            if len(full_res) > 0:
                res.extend(full_res)
        break
    return res

def GetC4DFiles(root_directory):
    filelist_path = os.path.join(root_directory,'file_list.txt')
    if os.path.exists(filelist_path):
        rvalue = gui.QuestionDialog('File list "' + filelist_path + '" was found. Load files from there?')
        if rvalue:
            FileList = []
            flp = open(filelist_path, 'r')
            for fline in flp:
                cur_line = fline.strip().decode('utf-8')
                if cur_line == '':
                    continue
                FileList.append(cur_line)
            flp.close()
        else:
            FileList = ListC4DFiles(root_directory)
    else:
        FileList = ListC4DFiles(root_directory)
    FileDict = {}
    for filepath in FileList:
        if FileDict.get(filepath, None) is None:
            if not os.path.exists(os.path.splitext(filepath)[0] + '.fbx'):
                FileDict.setdefault(filepath, True)
    FileList = FileDict.keys()
    FileList.sort()
    flp = open(filelist_path, 'w')
    for name in FileList:
        print >>flp, name.encode('utf-8')
    flp.close()
    return FileList


def printToLog(log_dir, message):
    filepath = os.path.join(log_dir,'result.txt')
    flog = open(filepath, 'a')
    print >>flog, message
    flog.close()

def ExportAllC4DFiles(root_directory):
    FilesToExport = GetC4DFiles(root_directory)
    total = len(FilesToExport)
    total_str = '%d'.format(total)
    total_length = len(total_str)
    printToLog(root_directory, ' --> STARTING EXPORTING {0:d} FILES IN "{1}"'.format(total, root_directory.encode('utf-8')))
    current = 0
    cur_succ = 0
    cur_fail = 0
    for path in FilesToExport:
        current += 1
        result = False
        message = ''
        documents.CloseAllDocuments()
        try:
            result,message = Export2FBX(path.encode('utf-8'), True)
        except Exception as e:
            message = str(e)
        if result:
            cur_succ += 1
        else:
            cur_fail += 1
        msg = '[ ' + repr(current).rjust(total_length) + ' / ' + repr(total).rjust(total_length) + '] '
        if cur_fail > 0:
            msg = '[ERRORS: ' + repr(cur_fail) + '] '
        msg += message + ' : "' + path.encode('utf-8') + '"'
        printToLog(root_directory, msg)
        documents.CloseAllDocuments()
    msg = ' <-- FINISH EXPORTING ' + repr(total) + ' FILES '
    if cur_fail > 0:
        msg += '(ERRORS: ' + repr(cur_fail) + ') '
    msg += ' IN "{0}"'.format(root_directory.encode('utf-8'))
    printToLog(root_directory, msg)

def StartConverting(root_directory=''):
    root_dir = os.path.join(SelectDirectory(root_directory),'')
    if os.path.isdir(root_dir):
        rvalue = gui.QuestionDialog('"' + root_dir + '" was selected. Start search for c4d files?')
    else:
        rvalue = gui.QuestionDialog('Path "' + root_dir + '" is not an existing directory. Continue anyway?')
    if rvalue:
        ExportAllC4DFiles(root_dir)
    gui.MessageDialog('Exiting from script')

class RenderFile(threading.C4DThread):
    
    def __init__(self, doc):
        self.dest_file = ''
        self.doc = doc.GetClone(c4d.COPYFLAGS_DOCUMENT)
        self.render_data = self.doc.GetActiveRenderData().GetData()
        self.bmp = bitmaps.BaseBitmap()
        self.bmp.Init(x=int(self.render_data[c4d.RDATA_XRES]), y=int(self.render_data[c4d.RDATA_YRES]), depth=24)
        self.status = False
        
    def Main(self):
        destfile = self.dest_file + '.jpg'
        self.status = documents.RenderDocument(self.doc, self.render_data, self.bmp, c4d.RENDERFLAGS_EXTERNAL)
        if self.status == c4d.RENDERRESULT_OK:
            self.status = self.bmp.Save(destfile, c4d.FILTER_JPG)
    
    def GetStatus(self):
        return self.status



class ExportFile(threading.C4DThread):
    
    def __init__(self, doc):
        self.dest_file = ''
        self.doc = doc.GetClone(c4d.COPYFLAGS_DOCUMENT)
        self.EXPORTER_ID = 1026370
        self.status = False
        
    def Main(self):
        destfile =  self.dest_file + '.fbx'
        self.status = documents.SaveDocument(self.doc, destfile, c4d.SAVEDOCUMENTFLAGS_DONTADDTORECENTLIST, self.EXPORTER_ID)
    
    def GetStatus(self):
        return self.status
    




def main():
    StartConverting()
    

if __name__=='__main__':
    main()



