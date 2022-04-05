Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        RichTextBox2.SelectionStart = 0
        Dim aranacaklar(), aranacak, arananlar() As String
        Dim sayi, i, j As Integer
        Dim bos As Boolean = False
        Dim ayni As Boolean = False

        ' inputla alınacak verinin boş olmaması için döngü kullandım
        Do While (bos = False)
            aranacak = InputBox("Aranacak Kelimeleri Giriniz")
            If aranacak = "" Then
                MsgBox("Lütfen Boş Bırakmayınız")
            Else
                bos = True
            End If
        Loop

        'aldığım veriyi boşluğa göre bölerek bir dizi içerisine aktardım
        'kullanıcı aynı kelimeyi 1'den fazla girmişse eğer fazlalıkları "" olarak değiştirdim
        aranacaklar = aranacak.Split(" ")
        Do While (i < aranacaklar.Length)
            j = i + 1
            Do While (j < aranacaklar.Length)
                If String.Compare(aranacaklar(i), aranacaklar(j)) = 0 Then
                    aranacaklar(i) = ""
                    j = aranacaklar.Length
                End If
                j += 1
            Loop
            i += 1
        Loop

        'richtextbox2 ye 1 i aktarardım ve ardından arananlar dizisine richtextbox2 den split yaptığım verileri ekledim
        RichTextBox2.Text = RichTextBox1.Text
        arananlar = RichTextBox2.Text.Split(" ")
        'arananların ilk elemanıyla aranacakların bütün elemanlarını sıra sıra kontrol ettim. 
        'arananların ilk elemanı bitince ikinci, üçüncü elemanına devam edecek şekilde algoritma kurdum
        'boolean değişkenini aynı değer bulup bulmadığını kontrol edebilmek için ekledim
        For i = 0 To arananlar.Length - 1
            ' alttaki for döngüsü bittiği zaman j yeniden 0 olarak işleme girebilmesi için j=0 ekledim
            j = 0
            ayni = False
            For j = 0 To aranacaklar.Length - 1
                ' yukarıda aynı olan değerleri "" ile değiştirmiştim, "" olan değerlerle boşa vakit kaybetmemesi için onlarla işlem yapmadım
                ' aranacaklar dizisindeki "" olmayan elemanlar karşılaştırmaya tutuldu
                If aranacaklar(j) <> "" Then
                    ' eğer arananların i. değeri ile aranacakların j. değeri aynıysa kullanacağım değişkenleri aldım
                    If String.Compare(arananlar(i), aranacaklar(j)) = 0 Then
                        sayi = RichTextBox2.Text.Trim().IndexOf(aranacaklar(j).Trim(), sayi)
                        RichTextBox2.SelectionLength = CInt(aranacaklar(j).Trim().Length)
                        ayni = True
                    End If
                End If
            Next
            ' boolean değişkenine göre işlem yaptırdım. Eğer yukarıda aynı kelimeyi bulduysa boolean True değere düştü
            ' boolean true değere düştüyse eğer renklendirme işlemlerini yaptırdım
            ' bir sonraki turda renklendirmeye en baştan başlamaması, zaten kontrolünü yapmış olduğu yerleri atlayıp başlaması için döngünün sonunda sayı değişkenini 
            ' kullanmış olduğum uzunluk kadar yükselttim. 
            If ayni = True Then
                RichTextBox2.SelectionStart = sayi
                RichTextBox2.SelectionColor = Color.White
                RichTextBox2.SelectionBackColor = Color.Red
                sayi += RichTextBox2.SelectionLength
            End If
        Next
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RichTextBox1.Text = ""
        RichTextBox2.Text = ""
    End Sub
End Class
