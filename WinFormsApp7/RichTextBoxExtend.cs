using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace RichTextBoxEx
{
    public class RichTextBoxWithShapes : RichTextBox
    {
        public void AddShape(Word.Shape shape)
        {
            // Tạo một InlineShape từ shape
            Word.InlineShape inlineShape = shape.ConvertToInlineShape();

            // Lấy hình ảnh của InlineShape và chuyển đổi nó thành một Bitmap
            Word.Range range = inlineShape.Range;
            range.CopyAsPicture();
            Image image = Clipboard.GetImage();

            // Chuyển đổi Bitmap thành một MemoryStream
            using (System.IO.MemoryStream stream = new System.IO.MemoryStream())
            {
                image.Save(stream, ImageFormat.Png);

                // Tạo một PictureBox để hiển thị hình ảnh
                PictureBox pictureBox = new PictureBox();
                pictureBox.Image = image;

                // Thêm PictureBox vào RichTextBox
                this.Controls.Add(pictureBox);
            }
        }
    }
}