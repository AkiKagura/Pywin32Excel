def export_source(application, path):
    for vbc in application.VBE.ActiveVBProject.VBComponents:
        # print(vbc.Name + str(vbc.Type))
        # i = workbook.VBProject.VBComponents(vbc.name).CodeModule.CountOfLines
        # if i >= 1:
        extend = ""
        if vbc.Type == 100:
            extend = ".cls"
        elif vbc.Type == 1:
            extend = ".bas"
        elif vbc.Type == 2:
            extend = ".cls"
        elif vbc.Type == 3:
            extend = ".frm"
        if extend != "":
            vbc.Export(path + "\\" + vbc.name + extend)

