# -*- coding: utf-8 -*-

import os
import sys
import sqlite3
import xlrd
import xlwt
from basic.common import getrootpath, makeDirsForFile, existFile
from basic.data import BIN2CHECK, VB_SESSION, Y_GRE_SESSION, SESSION_ORDER, SESSION_ORDER_NAME, SESSION_OVERLAP, LESSON, VB_LESSON, Y_GRE_LESSON, LESSON_ORDER, LESSON_COLUMN, VIDEO_LESSON, ROOM1103, ROOM1707

reload(sys)
sys.setdefaultencoding('utf-8');


def main():
    rootpath = getrootpath()
    inputpath = os.path.join(rootpath, 'Data')
    outputpath = os.path.join(rootpath, 'Results')

    # date = '20160630'
    date = raw_input('Please input a date (e.g.: 20160430): ')
    print '--'

    iPadContentsFile = os.path.join(inputpath, 'iPadContents.xls')
    iPadContentsFileX = os.path.join(inputpath, 'iPadContents.xlsx')
    appointmentFile = os.path.join(inputpath, 'Appointments-%s.xls' % date)
    appointmentFileX = os.path.join(inputpath, 'Appointments-%s.xlsx' % date)
    iPadAllocationFile = os.path.join(outputpath, 'iPadAllocation-%s.xls' % date)


    # Connect to SQLite database
    # database = ':memory:'
    database = 'appointments_%s.db' % date
    if not existFile(database, overwrite=True, displayInfo=False):
        conn = sqlite3.connect(database)
    c = conn.cursor()


    # Read iPad content data from iPadContents.xls(x) and write to appointments_yyyymmdd.db file
    try:
        c.executescript('''
            CREATE TABLE ipad_contents(
                ipad_label PRIMARY KEY,
                ipad_order,
                is_allocated,
                lesson_quantity,
                vb_intro,
                vb_1,
                vb_2,
                vb_3,
                vb_4,
                vb_5,
                vb_6,
                vb_7,
                vb_8,
                vb_9,
                gre_intro,
                gre_1,
                gre_2,
                gre_3,
                gre_4,
                gre_5,
                gre_6,
                gre_7,
                gre_8,
                gre_9,
                gre_test,
                aw_intro
            );
        ''')
    except Exception, e:
        print e

    try:
        excelFile = iPadContentsFile
        data = xlrd.open_workbook(excelFile)
    except Exception, e:
        excelFile = iPadContentsFileX
        data = xlrd.open_workbook(excelFile)
    print 'Import data: %s -> %s/ipad_contents' % (excelFile, database)
    table = data.sheet_by_index(0)
    headline = 1
    for row in range(table.nrows):
        if row >= headline:
            values = table.row_values(row)

            ipad_label = values[0]
            if ipad_label == u'Z00':
                continue
            vb_intro = int(values[1] == 1)
            vb_1 = int(values[2] == 1)
            vb_2 = int(values[3] == 1)
            vb_3 = int(values[4] == 1)
            vb_4 = int(values[5] == 1)
            vb_5 = int(values[6] == 1)
            vb_6 = int(values[7] == 1)
            vb_7 = int(values[8] == 1)
            vb_8 = int(values[9] == 1)
            vb_9 = int(values[10] == 1)
            gre_intro = int(values[11] == 1)
            gre_1 = int(values[12] == 1)
            gre_2 = int(values[13] == 1)
            gre_3 = int(values[14] == 1)
            gre_4 = int(values[15] == 1)
            gre_5 = int(values[16] == 1)
            gre_6 = int(values[17] == 1)
            gre_7 = int(values[18] == 1)
            gre_8 = int(values[19] == 1)
            gre_9 = int(values[20] == 1)
            gre_test = int(values[21] == 1)
            aw_intro = int(values[22] == 1)

            ipad_order = row - headline + 1
            is_allocated = 'N'
            lesson_quantity = vb_intro + vb_1 + vb_2 + vb_3 + vb_4 + vb_5 + vb_6 + vb_7 + vb_8 + vb_9 + gre_intro + gre_1 + gre_2 + gre_3 + gre_4 + gre_5 + gre_6 + gre_7 + gre_8 + gre_9 + gre_test + aw_intro

            values = (
                ipad_label,
                ipad_order,
                is_allocated,
                lesson_quantity,
                vb_intro,
                vb_1,
                vb_2,
                vb_3,
                vb_4,
                vb_5,
                vb_6,
                vb_7,
                vb_8,
                vb_9,
                gre_intro,
                gre_1,
                gre_2,
                gre_3,
                gre_4,
                gre_5,
                gre_6,
                gre_7,
                gre_8,
                gre_9,
                gre_test,
                aw_intro,
            )

            # Write data to *.db
            c.execute('INSERT INTO ipad_contents VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?);', values)

    # Save (commit) the changes
    conn.commit()
    print '--'


    # Read appointment data from Appointments-yyyymmdd.xls(x) and write to appointments_yyyymmdd.db file
    try:
        c.executescript('''
            CREATE TABLE appointments(
                appointment_uid PRIMARY KEY,
                name,
                lesson_type,
                lesson_status,
                video_status,
                session
            );
        ''')
    except Exception, e:
        print e

    try:
        excelFile = appointmentFile
        data = xlrd.open_workbook(excelFile)
    except Exception, e:
        excelFile = appointmentFileX
        data = xlrd.open_workbook(excelFile)
    print 'Import data: %s -> %s/appointments' % (excelFile, database)
    table = data.sheet_by_index(0)
    headline = 1
    appointment_uid = 0
    for row in range(table.nrows):
        if row >= headline:
            appointment_uid += 1
            values = table.row_values(row)

            name = values[0]
            lesson_status = values[1]
            video_status = values[3]
            session = values[2]

            # Obtain lesson type
            if lesson_status in VB_LESSON:
                lesson_type = u'VB'
            elif lesson_status in Y_GRE_LESSON:
                lesson_type = u'Y-GRE'
            else:
                print 'Invalid lesson status:', lesson_status

            # Validate lesson status
            try:
                lesson_status = VIDEO_LESSON[str(video_status)]
            except Exception, e:
                pass

            values = (
                appointment_uid,
                name,
                lesson_type,
                lesson_status,
                video_status,
                session,
            )

            # Write data to *.db
            c.execute('INSERT INTO appointments VALUES (?,?,?,?,?,?);', values)

    # Save (commit) the changes
    conn.commit()
    print '--'


    # Allocate iPads according to their contents and student appointments
    print 'Allocate iPads...'
    try:
        c.executescript('''
            CREATE TABLE allocation_results(
                allocation_uid PRIMARY KEY,
                name,
                lesson_type,
                lesson_status,
                lesson_order,
                video_status,
                session,
                session_order,
                ipad_label,
                supplementary_ipad_label
            );
        ''')
    except Exception, e:
        print e

    # Y-GRE Sessions
    iPadQuantity = []
    for lesson in Y_GRE_LESSON:
        c.execute('SELECT COUNT(ipad_label) FROM ipad_contents WHERE %s = 1;' % LESSON_COLUMN[lesson])
        iPadQuantity.append([lesson, c.fetchall()[0][0]])
    iPadQuantity.sort(key=lambda v: v[1])
    YGRELessonsSortedByiPadQuantity = [x[0] for x in iPadQuantity]

    allocation_uid = 0
    for session in Y_GRE_SESSION:
        allocatediPads = []
        for lesson in YGRELessonsSortedByiPadQuantity:
            c.execute('SELECT name, lesson_type, lesson_status, video_status FROM appointments WHERE session = ? AND lesson_status = ?;', (session, lesson, ))
            for name, lesson_type, lesson_status, video_status in c.fetchall():
                allocation_uid += 1
                c.execute('SELECT ipad_label FROM ipad_contents WHERE %s = 1 ORDER BY lesson_quantity ASC;' % LESSON_COLUMN[lesson_status])
                iPadCandidates = [x[0] for x in c.fetchall()]
                allocated = False
                for iPadCandidate in iPadCandidates:
                    if iPadCandidate not in allocatediPads:
                        allocated = True
                        allocatediPads.append(iPadCandidate)
                        c.execute('UPDATE ipad_contents SET is_allocated = "Y" WHERE ipad_label = ?', (iPadCandidate, ))
                        break
                if not allocated:
                    iPadCandidate = u'缺少资源'
                    print name, lesson_type, lesson_status, session, iPadCandidate

                lesson_order = LESSON_ORDER[lesson_status]
                session_order = SESSION_ORDER[session]

                iPadCandidate2 = u'N/A'

                values = (
                    allocation_uid,
                    name,
                    lesson_type,
                    lesson_status,
                    lesson_order,
                    video_status,
                    session,
                    session_order,
                    iPadCandidate,
                    iPadCandidate2,
                )

                # Write data to *.db
                c.execute('INSERT INTO allocation_results VALUES (?,?,?,?,?,?,?,?,?,?);', values)

    # Save (commit) the changes
    conn.commit()
    print '--'

    # VB Sessions
    iPadQuantity = []
    for lesson in VB_LESSON:
        c.execute('SELECT COUNT(ipad_label) FROM ipad_contents WHERE %s = 1;' % LESSON_COLUMN[lesson])
        iPadQuantity.append([lesson, c.fetchall()[0][0]])
    iPadQuantity.sort(key=lambda v: v[1])
    VBLessonsSortedByiPadQuantity = [x[0] for x in iPadQuantity]

    for session in VB_SESSION:
        allocatediPads = []
        for overlappedSession in SESSION_OVERLAP[session]:
            c.execute('SELECT ipad_label FROM allocation_results WHERE session_order = ?;', (SESSION_ORDER[overlappedSession], ))
            allocatediPads += [x[0] for x in c.fetchall()]
        YGREAllocatediPads = allocatediPads
        for lesson in VBLessonsSortedByiPadQuantity:
            c.execute('SELECT name, lesson_type, lesson_status, video_status FROM appointments WHERE session = ? AND lesson_status = ?;', (session, lesson, ))
            for name, lesson_type, lesson_status, video_status in c.fetchall():
                allocation_uid += 1
                c.execute('SELECT ipad_label FROM ipad_contents WHERE %s = 1 ORDER BY lesson_quantity ASC;' % LESSON_COLUMN[lesson_status])
                iPadCandidates = [x[0] for x in c.fetchall()]
                allocated = False
                for iPadCandidate in iPadCandidates:
                    if iPadCandidate not in allocatediPads:
                        allocated = True
                        allocatediPads.append(iPadCandidate)
                        c.execute('UPDATE ipad_contents SET is_allocated = "Y" WHERE ipad_label = ?;', (iPadCandidate, ))
                        break
                if not allocated:
                    iPadCandidate = u'缺少资源'
                    print name, lesson_type, lesson_status, session, iPadCandidate

                lesson_order = LESSON_ORDER[lesson_status]
                session_order = SESSION_ORDER[session]

                iPadCandidate2 = u'N/A'

                values = (
                    allocation_uid,
                    name,
                    lesson_type,
                    lesson_status,
                    lesson_order,
                    video_status,
                    session,
                    session_order,
                    iPadCandidate,
                    iPadCandidate2,
                )

                # Write data to *.db
                c.execute('INSERT INTO allocation_results VALUES (?,?,?,?,?,?,?,?,?,?);', values)

    # Save (commit) the changes
    conn.commit()
    print '--'


    # Excel cell formatting
    styleRed = xlwt.easyxf('font: color-index red')
    styleHighlight = xlwt.easyxf('pattern: pattern solid, fore_color gray25')
    styleNumber = xlwt.easyxf(num_format_str='#,##0.00')
    styleNumberRed = xlwt.easyxf('font: color-index red', num_format_str='#,##0.00')
    styleNumberHighlight = xlwt.easyxf('pattern: pattern solid, fore_color gray25', num_format_str='#,##0.00')
    stylePercentage = xlwt.easyxf(num_format_str='0.00%')
    stylePercentageRed = xlwt.easyxf('font: color-index red', num_format_str='0.00%')
    stylePercentageHighlight = xlwt.easyxf('pattern: pattern solid, fore_color gray25', num_format_str='0.00%')
    styleDate = xlwt.easyxf(num_format_str='yyyy/mm/dd')
    styleDateHighlight = xlwt.easyxf('pattern: pattern solid, fore_color gray25', num_format_str='yyyy/mm/dd')
    styleBold = xlwt.easyxf('font: bold on')
    styleBoldHighlight = xlwt.easyxf('font: bold on; pattern: pattern solid, fore_color gray25')


    # Write iPad allocation results to iPadAllocation-yyyymmdd.xls(x)
    excelFile = iPadAllocationFile
    print 'Export to %s' % excelFile
    wb = xlwt.Workbook()


    # Sheet: 11th Floor
    ws = wb.add_sheet(u'11楼')
    ws.write(0, 0, u'#', styleBold)
    ws.write(0, 1, u'预约时间段', styleBold)
    ws.write(0, 2, u'姓名', styleBold)
    ws.write(0, 3, u'课程进度', styleBold)
    ws.write(0, 4, u'视频进度', styleBold)
    ws.write(0, 5, u'机器编号', styleBold)

    row = 0
    c.execute('SELECT session, name, lesson_status, video_status, ipad_label FROM allocation_results ORDER BY session_order ASC, lesson_order ASC, ipad_label ASC;')
    for session, name, lesson_status, video_status, ipad_label in c.fetchall():
        if ipad_label in ROOM1103:
            row += 1
            ws.write(row, 0, row, styleBold)
            ws.write(row, 1, session)
            ws.write(row, 2, name)
            ws.write(row, 3, lesson_status)
            ws.write(row, 4, video_status)
            ws.write(row, 5, ipad_label)


    # Sheet: 17th Floor
    ws = wb.add_sheet(u'17楼')
    ws.write(0, 0, u'#', styleBold)
    ws.write(0, 1, u'预约时间段', styleBold)
    ws.write(0, 2, u'姓名', styleBold)
    ws.write(0, 3, u'课程进度', styleBold)
    ws.write(0, 4, u'视频进度', styleBold)
    ws.write(0, 5, u'机器编号', styleBold)

    row = 0
    c.execute('SELECT session, name, lesson_status, video_status, ipad_label FROM allocation_results ORDER BY session_order ASC, lesson_order ASC, ipad_label ASC;')
    for session, name, lesson_status, video_status, ipad_label in c.fetchall():
        if ipad_label in ROOM1707:
            row += 1
            ws.write(row, 0, row, styleBold)
            ws.write(row, 1, session)
            ws.write(row, 2, name)
            ws.write(row, 3, lesson_status)
            ws.write(row, 4, video_status)
            ws.write(row, 5, ipad_label)


    # Sheet: iPad Contents
    ws = wb.add_sheet(u'iPad资源统计')
    ws.write(0, 0, u'iPad# ', styleBold)
    ws.write(0, 1, u'分配状态', styleBold)
    ws.write(0, 2, u'VB总论', styleBold)
    ws.write(0, 3, u'L1', styleBold)
    ws.write(0, 4, u'L2', styleBold)
    ws.write(0, 5, u'L3', styleBold)
    ws.write(0, 6, u'L4', styleBold)
    ws.write(0, 7, u'L5', styleBold)
    ws.write(0, 8, u'L6', styleBold)
    ws.write(0, 9, u'L7', styleBold)
    ws.write(0, 10, u'L8', styleBold)
    ws.write(0, 11, u'L9', styleBold)
    ws.write(0, 12, u'GRE总论', styleBold)
    ws.write(0, 13, u'1st', styleBold)
    ws.write(0, 14, u'2nd', styleBold)
    ws.write(0, 15, u'3rd', styleBold)
    ws.write(0, 16, u'4th', styleBold)
    ws.write(0, 17, u'5th', styleBold)
    ws.write(0, 18, u'6th', styleBold)
    ws.write(0, 19, u'7th', styleBold)
    ws.write(0, 20, u'8th', styleBold)
    ws.write(0, 21, u'9th', styleBold)
    ws.write(0, 22, u'Test', styleBold)
    ws.write(0, 23, u'AW总论', styleBold)

    row = 0
    c.execute('SELECT ipad_label, is_allocated, vb_intro, vb_1, vb_2, vb_3, vb_4, vb_5, vb_6, vb_7, vb_8, vb_9, gre_intro, gre_1, gre_2, gre_3, gre_4, gre_5, gre_6, gre_7, gre_8, gre_9, gre_test, aw_intro FROM ipad_contents ORDER BY ipad_order ASC;')
    for ipad_label, is_allocated, vb_intro, vb_1, vb_2, vb_3, vb_4, vb_5, vb_6, vb_7, vb_8, vb_9, gre_intro, gre_1, gre_2, gre_3, gre_4, gre_5, gre_6, gre_7, gre_8, gre_9, gre_test, aw_intro in c.fetchall():
        row += 1
        ws.write(row, 0, ipad_label, styleBold)
        if is_allocated == u'Y':
            c.execute('SELECT session_order FROM allocation_results WHERE ipad_label = ?', (ipad_label, ))
            ws.write(row, 1, reduce(lambda x, y: x + u'、' + y, [SESSION_ORDER_NAME[x[0]] for x in c.fetchall()]), styleRed)
        else:
            ws.write(row, 1, u'未分配')
        ws.write(row, 2, BIN2CHECK[vb_intro])
        ws.write(row, 3, BIN2CHECK[vb_1])
        ws.write(row, 4, BIN2CHECK[vb_2])
        ws.write(row, 5, BIN2CHECK[vb_3])
        ws.write(row, 6, BIN2CHECK[vb_4])
        ws.write(row, 7, BIN2CHECK[vb_5])
        ws.write(row, 8, BIN2CHECK[vb_6])
        ws.write(row, 9, BIN2CHECK[vb_7])
        ws.write(row, 10, BIN2CHECK[vb_8])
        ws.write(row, 11, BIN2CHECK[vb_9])
        ws.write(row, 12, BIN2CHECK[gre_intro])
        ws.write(row, 13, BIN2CHECK[gre_1])
        ws.write(row, 14, BIN2CHECK[gre_2])
        ws.write(row, 15, BIN2CHECK[gre_3])
        ws.write(row, 16, BIN2CHECK[gre_4])
        ws.write(row, 17, BIN2CHECK[gre_5])
        ws.write(row, 18, BIN2CHECK[gre_6])
        ws.write(row, 19, BIN2CHECK[gre_7])
        ws.write(row, 20, BIN2CHECK[gre_8])
        ws.write(row, 21, BIN2CHECK[gre_9])
        ws.write(row, 22, BIN2CHECK[gre_test])
        ws.write(row, 23, BIN2CHECK[aw_intro])


    # Sheet: VB
    ws = wb.add_sheet(u'VB（签到表粘贴用数据）')
    ws.write(0, 0, u'#', styleBold)
    ws.write(0, 1, u'预约时间段', styleBold)
    ws.write(0, 2, u'预约姓名', styleBold)
    ws.write(0, 3, u'预约进度', styleBold)
    ws.write(0, 4, u'借出机号', styleBold)

    row = 0
    c.execute('SELECT session, name, lesson_status, video_status, ipad_label FROM allocation_results ORDER BY session_order ASC, ipad_label ASC, lesson_order ASC;')
    for session, name, lesson_status, video_status, ipad_label in c.fetchall():
        if ipad_label in ROOM1707:
            if lesson_status in VB_LESSON:
                row += 1
                ws.write(row, 0, row, styleBold)
                ws.write(row, 1, session)
                ws.write(row, 2, name)
                if video_status in ['', u'']:
                    ws.write(row, 3, lesson_status)
                else:
                    ws.write(row, 3, video_status)
                ws.write(row, 4, ipad_label)


    # Sheet: Intro
    ws = wb.add_sheet(u'总论（签到表粘贴用数据）')
    ws.write(0, 0, u'#', styleBold)
    ws.write(0, 1, u'预约时间段', styleBold)
    ws.write(0, 2, u'预约姓名', styleBold)
    ws.write(0, 3, u'预约进度', styleBold)
    ws.write(0, 4, u'借出机号', styleBold)

    row = 0
    c.execute('SELECT session, name, lesson_status, video_status, ipad_label FROM allocation_results ORDER BY session_order ASC, ipad_label ASC, lesson_order ASC;')
    for session, name, lesson_status, video_status, ipad_label in c.fetchall():
        if ipad_label in ROOM1103:
            row += 1
            ws.write(row, 0, row, styleBold)
            ws.write(row, 1, session)
            ws.write(row, 2, name)
            if video_status in ['', u'']:
                ws.write(row, 3, lesson_status)
            else:
                ws.write(row, 3, video_status)
            ws.write(row, 4, ipad_label)


    # Sheet: Y-GRE
    ws = wb.add_sheet(u'Y-GRE（签到表粘贴用数据）')
    ws.write(0, 0, u'#', styleBold)
    ws.write(0, 1, u'预约时间段', styleBold)
    ws.write(0, 2, u'预约姓名', styleBold)
    ws.write(0, 3, u'预约进度', styleBold)
    ws.write(0, 4, u'借出机号', styleBold)

    row = 0
    c.execute('SELECT session, name, lesson_status, video_status, ipad_label FROM allocation_results ORDER BY session_order ASC, ipad_label ASC, lesson_order ASC;')
    for session, name, lesson_status, video_status, ipad_label in c.fetchall():
        if ipad_label in ROOM1707:
            if lesson_status in Y_GRE_LESSON:
                row += 1
                ws.write(row, 0, row, styleBold)
                ws.write(row, 1, session)
                ws.write(row, 2, name)
                if video_status in ['', u'']:
                    ws.write(row, 3, lesson_status)
                else:
                    ws.write(row, 3, video_status)
                ws.write(row, 4, ipad_label)

    # Save file
    if not existFile(excelFile, overwrite=True):
        makeDirsForFile(excelFile)
        wb.save(excelFile)


    # We can also close the connection if we are done with it.
    # Just be sure any changes have been committed or they will be lost.
    conn.close()


if __name__ == '__main__':
    main()