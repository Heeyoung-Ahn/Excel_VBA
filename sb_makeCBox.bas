Attribute VB_Name = "sb_makeCBox"
Option Explicit

'------------------------------------------------------------------------------------------------------------------
'  �ܼ� �޺� ���� ��� �����:
'    - makeCBox(�޺��ڽ��̸�, Array("�׸�1", "�׸�2", ...), ListIndex��)
'    - ��: makecCBox(cbo1, Array("��ü", "���¿���", "���뿹��", "��Ÿ����", "����"), -1)
'------------------------------------------------------------------------------------------------------------------
Sub makeCBox(ByRef argCBox As MSForms.ComboBox, ByVal params As Variant, Optional ByVal index As Integer = -1)
    Dim cntParams As Integer
    
    For cntParams = LBound(params) To UBound(params)
        argCBox.AddItem params(cntParams)
    Next cntParams
    
    argCBox.ListIndex = index
End Sub
