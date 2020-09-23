Attribute VB_Name = "mdl_ID3"
Option Explicit

Public Type ID3Tag
    Header As String * 3
    SongTitle As String * 30
    Artist  As String * 30
    Album  As String * 30
    Year  As String * 4
    Comment As String * 30
    Genre  As Byte
End Type

Public Function GetID3Tag(ByVal FileName As String, Tag As ID3Tag) As Boolean

On Error GoTo GetID3TagError

Dim TempTag As ID3Tag
Dim FileNum As Long

    If Dir(FileName) = "" Then
        GetID3Tag = False
        Exit Function
    End If
    
    FileNum = FreeFile ' Get a handle
    
    Open FileName For Binary As FileNum
    Get FileNum, LOF(1) - 127, TempTag
    Close FileNum
    
    If TempTag.Header <> "TAG" Then
        GetID3Tag = False
    Else
        Tag = TempTag
        GetID3Tag = True
    End If

    Close FileNum
    Exit Function

GetID3TagError:
    
    Close FileNum
    GetID3Tag = False
    
End Function
Public Function SetID3Tag(ByVal FileName As String, Tag As ID3Tag) As Boolean

On Error GoTo SetID3TagError

Dim FileNum As Long

    If Dir(FileName) = "" Then
        SetID3Tag = False
        Exit Function
    End If
    
    Tag.Header = "TAG"
    
    FileNum = FreeFile
    
    Open FileName For Binary As FileNum
    Put FileNum, LOF(1) - 127, Tag
    Close FileNum
    
    SetID3Tag = True
    Close FileNum
    
    Exit Function

SetID3TagError:
    
    Close FileNum
    SetID3Tag = False
    
End Function

Public Function SetID3TagDirect(ByVal FileName As String, _
        ByVal Artist_30 As String, ByVal SongTitle_30 As String, _
        ByVal Album_30 As String, ByVal Comment_30 As String, _
        ByVal Year_4 As String, ByVal Genre_B255 As Byte) As Boolean

Dim Tag As ID3Tag

On Error GoTo SetID3TagDirectError

Dim FileNum As Long

    If Dir(FileName) = "" Then
        SetID3TagDirect = False
        Exit Function
    End If
    
    Tag.Header = "TAG"
    Tag.Artist = Artist_30
    Tag.SongTitle = SongTitle_30
    Tag.Album = Album_30
    Tag.Comment = Comment_30
    Tag.Year = Year_4
    Tag.Genre = Genre_B255
    
    FileNum = FreeFile
    
    Open FileName For Binary As FileNum
    Put FileNum, LOF(1) - 127, Tag
    Close FileNum
    
    SetID3TagDirect = True
    Close FileNum
    
    Exit Function

SetID3TagDirectError:

    Close FileNum
    SetID3TagDirect = False
    
End Function




