import excelChecker
if __name__ == '__main__':
    # creating excel checker object:
    ec1 = excelChecker.ExcelChecker()

    cs_cdcs = ['CS F211', 'CS F241', 'CS F212', 'BITS F225']

    # defining cdcs array
    cdcs = ['ECON F341', 'ECON F342', 'ECON F343']
    # edit your cdcs here
    # cdcs = ['ECON F211', 'PHY F241', 'PHY F242', 'PHY F243', 'PHY F244']
    # 'CS F211', 'CS F241', 'CS F212', 'BITS F225'  -> CS CDCs
    # 'MATH F112', 'MATH F113', 'ECON F211'
    # 'BITS F463', 'CS F241' -> check for cdc clash, works successfully
    # 'BITS F110', 'BITS F111'
    # 'CHEM F111'
    # 'CS F211', 'CS F241', 'CS F212', 'ECON F341', 'ECON F342', 'ECON F343' -> Msc Eco + CS CDCs

    ec1.mainchecker(cdcs, 'TIMETABLE - II SEMESTER 2024 -25_removed.xlsx')
    ec1.mainchecker(cs_cdcs, 'DRAFT TIMETABLE (2).xlsx')

