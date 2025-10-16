Attribute VB_Name = "Module1"
Sub VBA_separate_table_to_files()

    Application.DisplayAlerts = False '��������� ���������� �� �������� �����
    
    '��������� ������ �� �������� �����
    Dim sourceWorkbook As Workbook
    Set sourceWorkbook = ActiveWorkbook
    
    '������ ����� ������������ ������� - ����� ������� �� ��������� �����

    For Each cell In Range("ID") '� ������� ��������� ����������� ������������� �������
        Range("�����").AutoFilter Field:=8, Criteria1:=cell.Value '� ������� ��������� ������������ ������������� �������
        Range("�����[#All]").SpecialCells(xlCellTypeVisible).Copy
        Sheets.Add
        ActiveSheet.Paste
        ActiveSheet.Name = cell.Value
        ActiveSheet.UsedRange.Columns.AutoFit
    Next cell

'������ ����� ������������ ������� - ����� ����� �� ��������� �����

    Dim s As Worksheet
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    For Each s In wb.Worksheets
        '���������� ���� "������" ��� �������� ��������� ������
        If s.Name <> "������" Then
            s.Copy                                                  '��������� ���� ��� ����� ����
            ActiveWorkbook.SaveAs wb.Path & "\" & s.Name & ".xlsx"  '��������� ����
            ActiveWorkbook.Close                                    '��������� ��� ��������� �����
        End If
    Next s

'������ ����� ������������ ������� - ������� ����� �� ������� �����, ����� ���������

    Dim mySheet As Worksheet
    For Each mySheet In sourceWorkbook.Worksheets
        If mySheet.Name <> "������" Then
            mySheet.Delete
        End If
    Next mySheet

    Application.DisplayAlerts = True '�������� ���������� �� �������� ����� �������
    
    '��������� �������� ���� ��� ����������
    sourceWorkbook.Close SaveChanges:=False

End Sub
