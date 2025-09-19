Attribute VB_Name = "M35_Models"
'============================================================
' M35_Models : �p�[�X�E����Ɏg���f�[�^���f���iUDT/Enum�j
'============================================================
Option Explicit

' �ٓ��̎�ʁi�������蕪�ށj
Public Enum MovementKind
    mvUnknown = 0
    mvBorn = 1          ' �o��
    mvPurchaseIn = 2    ' ����/�]��/�w��/���q
    mvSaleOut = 3       ' ���p/�]�o/�o��
    mvDeath = 4         ' ���S
    mvTransfer = 5      ' ���̑��̈ړ��i�{�ݓ�/�{�݊ԂȂǁj
End Enum

' �ٓ�1���Ԃ�
Public Type MovementRow
    No As String                ' �ԍ��i�T�C�g�̘A�ԂȂǁj
    Kind As MovementKind        ' ��ʁi���Enum�j
    DateValue As Date           ' �ٓ��N����
    Pref As String              ' ���{�{�ݓs���{��
    City As String              ' ���{�{�ݎs�撬��
    Name As String              ' �����܂��͖��́i�q�ꖼ�j
    RawKind As String           ' ���́u�ٓ����e�v������i�ۑS�j
End Type

' �̃v���t�B�[��
Public Type CowProfile
    CowID As String                 ' �̎��ʔԍ��i10���j
    BirthDate As Date               ' �o���̔N����
    Sex As String                   ' ���Y�̕�
    MotherID As String              ' �ꋍ�̌̎��ʔԍ�
    Breed As String                 ' ���
End Type

' ���茋�ʁi�̂��Ɓj
Public Type CowRecord
    Profile As CowProfile
    Movements() As MovementRow

    ' ����t���O
    BornAtFarm As Boolean
    AtStart As Boolean
    BornInTerm As Boolean
    PurchasedInTerm As Boolean
    SoldInTerm As Boolean
    DiedInTerm As Boolean
    AtEnd As Boolean

    ' ��\���t�E�����i���[�p�ɕ֗��j
    PurchaseDate As Variant ' Date or Empty
    PurchaseFrom As String
    SaleDate As Variant
    SaleTo As String
    DeathDate As Variant
End Type

