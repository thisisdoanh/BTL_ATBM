using System.Security.Cryptography;
using System.Text;
using System.Numerics;
using System.Windows.Forms;
using Spire.Doc.Documents;
using Spire.Doc;
using System.ComponentModel;
using System.Xml.Linq;
using Spire.Doc.Fields;
using Microsoft.Office.Interop.Word;
using Document = Spire.Doc.Document;
using Section = Spire.Doc.Section;
using Paragraph = Spire.Doc.Documents.Paragraph;
using Font = System.Drawing.Font;
using Color = System.Drawing.Color;
using Application = Microsoft.Office.Interop.Word.Application;
using RichTextBoxEx;
using System;

namespace WinFormsApp7
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        RSA rsa = null;
        Boolean openFileWord = false;


        long[] listPrime = { 1000151, 1000159, 1000171, 1000183, 1000187,
            1000193, 1000199, 1000211, 1000213, 1000231, 1000249, 1000253,
            1000273, 1000289, 1000291, 1000303, 1000313, 1000333, 1018309,
            1018313, 1018337, 1018357, 1018411, 1023107, 1023133, 1023163,
            1023167 , 1036331, 1036339, 1036349, 1036351, 1036363, 1036367,
            1036369, 1058507, 1058543, 1058549, 1058567, 1058591, 1058593,
            1099313, 1099327, 1099337, 1099363, 1099369, 1099391, 1099393,
            1099409, 1221707, 1221749, 1221751, 1221761, 1221767, 1221791,
            1232071, 1232083, 1232089, 1232171, 1232183, 1232201, 1232213,
            1248353, 1248383, 1248391, 1248407, 1248413, 1248427, 1248449,
            1270513, 1270531, 1270537, 1270541, 1270547, 1270559, 1270561,
            3224789, 3224791, 3224797, 3224801, 3224833, 3224857, 3224861,
            3236531, 3236539, 3236543, 3236557, 3236587, 3236591, 3236621,
            3281081, 3281087, 3281093, 3281101, 3281137, 3281141, 3281149,
            3299617, 3299633, 3299641, 3299651, 3299677, 3299687, 3299689,
            8726573, 8726587, 8726611, 8726633, 8726647, 8726671, 8726677,
            8745311, 8745329, 8745343, 8745353, 8745367, 8745419, 8745437,
            8778593, 8778607, 8778629, 8778641, 8778643, 8778667, 8778701,
            8799887, 8799899, 8799911, 8799913, 8799919, 8799929, 8799941,
            9800009, 9800027, 9800053, 9800101, 9800113, 9800129, 9800137,
            9821419, 9821429, 9821443, 9821453, 9821489, 9821503, 9821507,
            9833491, 9833519, 9833533, 9833561, 9833597, 9833609, 9833617,
            9833623, 9833641, 9833647, 9833671, 9833687, 9833689, 9833711,
            9833717, 9833729, 9841441, 9841457, 9841483, 9841537, 9841543,
            9841561, 9841591, 9841597, 9841607, 9841609, 9841619, 9841627,
            9841661, 9841693, 9841703, 9841721, 9898433, 9898451, 9898457,
            9898481, 9898519, 9898543, 9898547, 9943253, 9943301, 9943303,
            9943319, 9943331, 9943333, 9943357, 9960347, 9960359, 9960367,
            9960371, 9960383, 9960407, 9960443, 9979639, 9979643, 9979681,
            9979727, 9979759, 9979763, 9979771, 9981731, 9981733, 9981739,
            9981779, 9981787, 9981799, 9981809, 9981847, 9981863, 9981869,
            9981889, 9981899, 9981911, 9981913, 9981919, 9981929, 9999883,
            9999889, 9999901, 9999907, 9999929, 9999931, 9999937, 9999943,
            9999971, 9999973, 9999991};

        private void btnRandomPQ_Click(object sender, EventArgs e)
        {
            Random rd = new Random();
            this.txtB.Text = "";
            this.txtA.Text = "";
            this.txtBanRoCheck.Clear();
            this.txtChuKyCheck.Clear();
            this.txtThongBao.Clear();
            this.txtP.Text = listPrime[rd.Next(0, listPrime.Length)].ToString();
            this.txtQ.Text = listPrime[rd.Next(0, listPrime.Length)].ToString();
            while (this.txtP.Text.Equals(this.txtQ.Text))
            {
                this.txtQ.Text = listPrime[rd.Next(0, listPrime.Length)].ToString();
            }

        }

        bool independencePQ(BigInteger p, BigInteger q)
        {
            return Convert.ToBoolean(p.Equals(q)) ? false : true;
        }

        ///Miller-Rabin ktra snt lớn
        public int[] getA(BigInteger n)
        {
            int[] a;
            if (BigInteger.Compare(n, 2047) < 0)
            {
                a = new int[1] { 2 };
                return a;
            }
            else
            {
                if (BigInteger.Compare(n, 1373653) < 0)
                {
                    a = new int[2] { 2, 3 };
                    return a;
                }
                else
                {
                    if (BigInteger.Compare(n, 9080191) < 0)
                    {
                        a = new int[2] { 31, 73 };
                        return a;
                    }
                    else
                    {
                        if (BigInteger.Compare(n, 4759123141) < 0)
                        {
                            a = new int[3] { 2, 7, 61 };
                            return a;
                        }
                        else
                        {
                            if (BigInteger.Compare(n, BigInteger.Parse("1122004669633")) < 0)
                            {
                                a = new int[4] { 2, 13, 23, 1662803 };
                                return a;
                            }
                            else
                            {
                                if (BigInteger.Compare(n, BigInteger.Parse("2152302898747")) < 0)
                                {
                                    a = new int[5] { 2, 3, 5, 7, 11 };
                                    return a;
                                }
                                else
                                {
                                    if (BigInteger.Compare(n, BigInteger.Parse("3474749660383")) < 0)
                                    {
                                        a = new int[6] { 2, 3, 5, 7, 11, 13 };
                                        return a;
                                    }
                                    else
                                    {
                                        if (BigInteger.Compare(n, BigInteger.Parse("341550071728321")) < 0)
                                        {
                                            a = new int[7] { 2, 3, 5, 7, 11, 13, 17 };
                                            return a;
                                        }
                                        else
                                        {
                                            if (BigInteger.Compare(n, BigInteger.Parse("3825123056546413051")) < 0)
                                            {
                                                a = new int[9] { 2, 3, 5, 7, 11, 13, 17, 19, 23 };
                                                return a;
                                            }
                                            else
                                            {
                                                if (BigInteger.Compare(n, BigInteger.Parse("318665857834031151167461")) < 0)
                                                {
                                                    a = new int[12] { 2, 3, 5, 7, 11, 13, 17, 19, 23, 29, 31, 37 };
                                                    return a;
                                                }
                                                else
                                                {
                                                    if (BigInteger.Compare(n, BigInteger.Parse("3317044064679887385961981")) < 0)
                                                    {
                                                        a = new int[13] { 2, 3, 5, 7, 11, 13, 17, 19, 23, 29, 31, 37, 41 };
                                                        return a;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return a = new int[1] { 0 };
        }

        public bool millerRabinTesting(BigInteger n)
        {
            if (BigInteger.Compare(n, 2) < 0) return false;
            if (n.Equals(2) || n.Equals(3)) return true;
            int[] a = getA(n);
            BigInteger d = BigInteger.Subtract(n, 1);
            BigInteger s = decompose(ref d);
            for (int i = 0; i < a.Length; i++)
            {
                BigInteger p = modPower(a[i], d, n);
                if (p.Equals(1))
                {
                    return true;
                }
                while (s > 0)
                {
                    if (p.Equals(BigInteger.Subtract(n, 1))) return true;
                    p = BigInteger.Remainder(BigInteger.Multiply(p, p), n);
                    s--;
                }
                return false;
            }
            return false;
        }

        public BigInteger decompose(ref BigInteger p)
        {
            BigInteger i = 0;
            while (BigInteger.Remainder(p, 2).Equals(0))
            {
                i++;
                p /= 2;
            }
            return i;
        }

        public BigInteger modPower(BigInteger a, BigInteger b, BigInteger p)
        {
            if (b.Equals(1))
                return a;
            else
            {
                BigInteger x = modPower(a, BigInteger.Divide(b, 2), p);

                if (BigInteger.Remainder(b, 2).Equals(0))
                {
                    return BigInteger.Remainder(BigInteger.Multiply(x, x), p);
                }
                else
                {
                    return BigInteger.Remainder(BigInteger.Multiply(BigInteger.Multiply(x, x), a), p);
                }
            }
        }

        private void btnTaoKhoa_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.txtP.Text.Trim().Equals(""))
                {
                    this.txtA.Text = "";
                    this.txtB.Text = "";
                    MessageBox.Show("Bạn chưa nhập giá trị của P", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (this.txtQ.Text.Trim().Equals(""))
                {
                    this.txtA.Text = "";
                    this.txtB.Text = "";
                    MessageBox.Show("Bạn chưa nhập giá trị của Q", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                BigInteger p = 0, q = 0;

                try
                {
                    p = BigInteger.Parse(this.txtP.Text.Trim());
                }
                catch (Exception)
                {
                    this.txtA.Text = "";
                    this.txtB.Text = "";
                    MessageBox.Show("P không hợp lệ. Mời nhập lại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                try
                {
                    q = BigInteger.Parse(this.txtQ.Text.Trim());
                }
                catch (Exception)
                {
                    this.txtA.Text = "";
                    this.txtB.Text = "";
                    MessageBox.Show("Q không hợp lệ. Mời nhập lại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }


                if (!independencePQ(p, q))
                {
                    this.txtA.Text = "";
                    this.txtB.Text = "";
                    MessageBox.Show("P, Q phải là 2 giá trị độc lập", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else
                {
                    if (!millerRabinTesting(p))
                    {
                        this.txtA.Text = "";
                        this.txtB.Text = "";
                        MessageBox.Show("P không phải số nguyên tố", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else
                    {
                        if (!millerRabinTesting(q))
                        {
                            this.txtA.Text = "";
                            this.txtB.Text = "";
                            MessageBox.Show("Q không phải số nguyên tố", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        else
                        {
                            if (BigInteger.Multiply(p, q) < 256)
                            {
                                this.txtA.Text = "";
                                this.txtB.Text = "";
                                MessageBox.Show("Hãy chọn 2 số nguyên tố lớn hơn để tăng tính bảo mật", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                            else
                            {
                                rsa = new RSA(p, q);
                                this.txtB.Text = rsa.B.ToString();
                                this.txtA.Text = rsa.A.ToString();

                            }
                        }


                    }

                }

            }
            catch (FormatException)
            {
                this.txtP.ResetText();
                this.txtQ.ResetText();
                MessageBox.Show("P, Q không hợp lệ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }

        //Function tai file len
        public string upFile()
        {
            ///Mowr hộp thoại đeer duyệt thư mục
            OpenFileDialog openFile = new OpenFileDialog();

            ///lọc các file có định dạng text
            openFile.Filter = "|*.txt";

            if (openFile.ShowDialog() == DialogResult.OK)
            {
                MessageBox.Show($"Bạn đã mở thành công file sau: {openFile.FileName}", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                string text = File.ReadAllText(openFile.FileName);
                return text;
            }
            return "";
        }


        //Function click button tải file bản rõ
        private void btnFileBanRo_Click(object sender, EventArgs e)
        {
            this.txtBanRo.Text = upFile();
            if (this.txtBanRo.Text.Trim().Equals(""))
            {
                MessageBox.Show("Chưa có văn bản nào được nhập vào!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }


        //Đưa chuỗi text vào hàm băm MD5

        public string hashMD5(string text)
        {
            ///convert string về byte
            byte[] textByte = Encoding.ASCII.GetBytes(text);
            MD5 md5 = MD5.Create();
            byte[] hash = md5.ComputeHash(textByte);
            StringBuilder hashSB = new StringBuilder();
            foreach (byte b in hash)
            {
                hashSB.Append(b.ToString("X2"));
            }
            //MessageBox.Show(hashSB.ToString());
            return hashSB.ToString();
        }
        /*
        public string hashMD5(string input)
        {
            uint[] t = new uint[64];
            for (int i = 0; i < 64; i++)
            {
                t[i] = (uint)(4294967296L * Math.Abs(Math.Sin(i + 1)));
            }

            byte[] message = Encoding.ASCII.GetBytes(input);
            uint messageLength = (uint)message.Length * 8;

            uint[] x = new uint[16];
            for (int i = 0; i < message.Length; i++)
            {
                int j = i % 4;
                int k = i / 4;
                x[k] |= (uint)(message[i] << (j * 8));
            }

            x[(int)(messageLength / 32) % 16] |= (uint)(1 << (int)(messageLength % 32));

            x[14] = messageLength;

            uint a = 0x67452301;
            uint b = 0xefcdab89;
            uint c = 0x98badcfe;
            uint d = 0x10325476;

            for (int i = 0; i < 64; i++)
            {
                uint f, g;
                if (i < 16)
                {
                    f = (b & c) | ((~b) & d);
                    g = (uint)i;
                }
                else if (i < 32)
                {
                    f = (d & b) | ((~d) & c);
                    g = (uint)(5 * i + 1) % 16;
                }
                else if (i < 48)
                {
                    f = b ^ c ^ d;
                    g = (uint)(3 * i + 5) % 16;
                }
                else
                {
                    f = c ^ (b | (~d));
                    g = (uint)(7 * i) % 16;
                }

                uint temp = d;
                d = c;
                c = b;
                b = b + RotateLeft((a + f + x[g] + t[i]), (int)Math.Pow(2, (i % 4) * 2 + 1));
                a = temp;
            }

            return string.Format("{0:x8}{1:x8}{2:x8}{3:x8}", a, b, c, d);
            
        }
        
        private uint RotateLeft(uint value, int count)
        {
            return (value << count) | (value >> (32 - count));
        }

        */
        public string hashMD5Word(RichTextBox rtb)
        {
            // Chuyển text và format qua html
            string html = "<html><body>" + rtb.Rtf + "</body></html>";

            // Convert html sang bytes
            byte[] htmlBytes = Encoding.UTF8.GetBytes(html);

            // băm
            using (MD5 md5 = MD5.Create())
            {
                byte[] hashBytes = md5.ComputeHash(htmlBytes);

                // Convert the hash to a hex string.
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < hashBytes.Length; i++)
                {
                    sb.Append(hashBytes[i].ToString("X2"));
                }

                return sb.ToString();
            }
        }

        public string hashMD5WWord1(RichTextBox rtb)
        {
            // Chuyển text và format qua html
            string html = "<html><body>" + rtb.Rtf + "</body></html>";

            // Convert html sang bytes
            byte[] htmlBytes = Encoding.UTF8.GetBytes(html);

            byte[] hash = ComputeMD5(htmlBytes);
            StringBuilder sb = new StringBuilder();

            foreach (byte b in hash)
            {
                sb.Append(b.ToString("x2"));
            }
            return sb.ToString();
        }
        public byte[] ComputeMD5(byte[] input)
        {
            uint[] s = new uint[] { 7, 12, 17, 22, 5, 9, 14, 20, 4, 11, 16, 23, 6, 10, 15, 21 };
            uint[] k = new uint[]
            {
                0xd76aa478, 0xe8c7b756, 0x242070db, 0xc1bdceee,
                0xf57c0faf, 0x4787c62a, 0xa8304613, 0xfd469501,
                0x698098d8, 0x8b44f7af, 0xffff5bb1, 0x895cd7be,
                0x6b901122, 0xfd987193, 0xa679438e, 0x49b40821,
                0xf61e2562, 0xc040b340, 0x265e5a51, 0xe9b6c7aa,
                0xd62f105d, 0x02441453, 0xd8a1e681, 0xe7d3fbc8,
                0x21e1cde6, 0xc33707d6, 0xf4d50d87, 0x455a14ed,
                0xa9e3e905, 0xfcefa3f8, 0x676f02d9, 0x8d2a4c8a,
                0xfffa3942, 0x8771f681, 0x6d9d6122, 0xfde5380c,
                0xa4beea44, 0x4bdecfa9, 0xf6bb4b60, 0xbebfbc70,
                0x289b7ec6, 0xeaa127fa, 0xd4ef3085, 0x04881d05,
                0xd9d4d039, 0xe6db99e5, 0x1fa27cf8, 0xc4ac5665,
                0xf4292244, 0x432aff97, 0xab9423a7, 0xfc93a039,
                0x655b59c3, 0x8f0ccc92, 0xffeff47d, 0x85845dd1,
                0x6fa87e4f, 0xfe2ce6e0, 0xa3014314, 0x4e0811a1,
                0xf7537e82, 0xbd3af235, 0x2ad7d2bb, 0xeb86d391
                };

            uint[] a = new uint[] { 0x67452301, 0xefcdab89, 0x98badcfe, 0x10325476 };

            byte[] paddedInput = Pad(input);
            int n = paddedInput.Length / 64;

            for (int i = 0; i < n; i++)
            {
                uint[] m = new uint[16];
                for (int j = 0; j < 16; j++)
                {
                    int index = i * 64 + j * 4;
                    m[j] = (uint)paddedInput[index] | ((uint)paddedInput[index + 1] << 8) |
                        ((uint)paddedInput[index + 2] << 16) | ((uint)paddedInput[index + 3] << 24);
                }

                uint a0 = a[0], b0 = a[1], c0 = a[2], d0 = a[3];

                for (int j = 0; j < 64; j++)
                {
                    uint f, g;
                    if (j < 16)
                    {
                        f = (b0 & c0) | (~b0 & d0);
                        g = (uint)j;
                    }
                    else if (j < 32)
                    {
                        f = (d0 & b0) | (~d0 & c0);
                        g = (uint)(5 * j + 1) % 16;
                    }
                    else if (j < 48)
                    {
                        f = b0 ^ c0 ^ d0;
                        g = (uint)(3 * j + 5) % 16;
                    }
                    else
                    {
                        f = c0 ^ (b0 | ~d0);
                        g = (uint)(7 * j) % 16;
                    }

                    uint temp = d0;
                    d0 = c0;
                    c0 = b0;
                    b0 = (uint)(b0 + ((int)(a0 + f + k[j] + m[g]) << (int)s[j % 4]));
                    a0 = temp;
                }

                a[0] += a0;
                a[1] += b0;
                a[2] += c0;
                a[3] += d0;
            }

            byte[] result = new byte[16];
            int pos = 0;
            for (int i = 0; i < 4; i++)
            {
                uint value = a[i];
                result[pos++] = (byte)(value & 0xff);
                result[pos++] = (byte)((value >> 8) & 0xff);
                result[pos++] = (byte)((value >> 16) & 0xff);
                result[pos++] = (byte)((value >> 24) & 0xff);
            }

            return result;
        }

        // Hàm này dùng để bổ sung các byte để đảm bảo độ dài của mảng byte đầu vào là bội số của 64
        private byte[] Pad(byte[] input)
        {
            int length = input.Length;
            int padding = 64 - ((length + 8) % 64);
            int newLength = length + padding + 8;
            byte[] result = new byte[newLength];

            for (int i = 0; i < length; i++)
            {
                result[i] = input[i];
            }

            result[length] = 0x80;

            for (int i = length + padding + 1; i < newLength; i++)
            {
                result[i] = 0;
            }

            byte[] lengthBytes = BitConverter.GetBytes((ulong)length * 8);
            for (int i = 0; i < 8; i++)
            {
                result[newLength - 8 + i] = lengthBytes[i];
            }

            return result;
        }
        public string toBinary(BigInteger a, int n)
        {
            StringBuilder sb = new StringBuilder();
            while (a > 0)
            {
                sb.Insert(0, a % n);
                a /= n;
            }
            return sb.ToString();
        }

        public string thuatToanBinhPhuongVaNhan(string ch, BigInteger m, BigInteger b)
        {
            BigInteger c = 0;
            for (int i = 0; i < ch.Length; i++)
            {
                c = BigInteger.Add(c, (int)ch[i]);
            }
            //convert số mũ qua dạng nhị phân
            string charNhiPhan = toBinary(b, 2);
            char[] mangNhiPhan = charNhiPhan.ToCharArray();
            BigInteger p = 1;
            foreach (char item in mangNhiPhan)
            {
                p = BigInteger.Pow(p, 2);
                p = BigInteger.Remainder(p, m);

                if (item == '1')
                {
                    p = BigInteger.Multiply(p, c);
                    p = BigInteger.Remainder(p, m);
                }
            }

            return p.ToString();
        }

        //Function mã hoá RSA
        public string maHoaRSA(string[] textMaHoa, BigInteger a, BigInteger b)
        {
            string result = "";
            foreach (string item in textMaHoa)
            {
                BigInteger c = BigInteger.Parse(thuatToanBinhPhuongVaNhan(item, a, b));
                result = result + c.ToString() + "-";
            }
            result = result.Remove(result.Length - 1, 1);
            return result;
        }

        string txtHashMD5 = "";
        private void btnKy_Click(object sender, EventArgs e)
        {
            //Xét có văn bản hay chưa
            if (this.txtBanRo.Text.Trim().Equals(""))
            {
                MessageBox.Show("Vui lòng nhập thông điệp!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (this.txtP.Text.Trim().Equals(""))
            {
                MessageBox.Show("Vui lòng nhập số nguyên tố P!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (this.txtQ.Text.Trim().Equals(""))
            {
                MessageBox.Show("Vui lòng nhập số nguyên tố Q!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (this.txtB.Text.Trim().Equals(""))
            {
                MessageBox.Show("Vui lòng bấm nút \"Tạo khoá\" trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (BigInteger.Compare(BigInteger.Multiply(BigInteger.Parse(this.txtP.Text.Trim()), BigInteger.Parse(this.txtQ.Text.Trim())), rsa.N) != 0)
            {
                MessageBox.Show("P hoặc Q đã bị thay đổi, vui lòng bấm nút \"Tạo khoá\" rồi thử lại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //Băm
            string text = this.txtBanRo.Text;
            string hashMD5Text = "";
            if (openFileWord == true)
            {
                hashMD5Text = hashMD5Word(this.txtBanRo).ToUpper();

            }
            else
            {
                hashMD5Text = hashMD5(this.txtBanRo.Text).ToUpper();
            }
            txtHashMD5 = hashMD5Text;
            // MessageBox.Show(txtHashMD5);

            //Ký
            string maHoa = hashMD5Text;
            string[] charMaHoa = new string[32];

            int lenMaHoa = maHoa.Length;
            for (int i = 0; i < lenMaHoa; i++)
            {
                charMaHoa[i] = maHoa.Substring(i, 1);
            }

            BigInteger a = BigInteger.Parse(this.txtA.Text);
            BigInteger b = BigInteger.Multiply(rsa.p, rsa.q);
            this.txtChuKy.Text = maHoaRSA(charMaHoa, b, a);
            if (openFileWord)
                openFileWord = false;
        }


        private void btnChuyen_Click(object sender, EventArgs e)
        {
            if (this.txtChuKy.Text.Equals(""))
            {
                MessageBox.Show("Chưa có chữ ký!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            this.txtBanRoCheck.Clear();
            this.txtBanRo.SelectAll();
            this.txtBanRo.Copy();
            this.txtBanRo.SelectAll();
            this.txtBanRo.Copy();
            this.txtBanRoCheck.Paste();
            this.txtChuKyCheck.Text = this.txtChuKy.Text;
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            ///check đã có văn bản được ký chưa
            if (this.txtChuKy.Text.Equals(""))
            {
                MessageBox.Show("Chưa có văn bản nào được ký!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using (SaveFileDialog saveFile = new SaveFileDialog())
            {
                saveFile.Filter = "Text files (*.txt)|*.txt";
                if (saveFile.ShowDialog() == DialogResult.OK)
                {
                    StreamWriter writer = new StreamWriter(saveFile.FileName);
                    writer.Write(this.txtChuKy.Text.ToString());
                    writer.Close();
                    MessageBox.Show("Chữ ký đã được lưu tại " + saveFile.FileName, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }

        }


        private void btnLuufileWord_click(object sender, EventArgs e)
        {

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Word documents (*.docx)|*.docx";
            saveFileDialog.Title = "Save a Word Document";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filename = saveFileDialog.FileName;

                // Khởi động Microsoft Word
                Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document doc = word.Documents.Add();

                // Thêm nội dung vào tài liệu Word
                Microsoft.Office.Interop.Word.Paragraph paragraph = doc.Content.Paragraphs.Add();
                paragraph.Range.Text = this.txtChuKy.Text;
                paragraph.Range.InsertParagraphAfter();

                // Lưu tài liệu Word dưới định dạng .docx
                object fileName = filename;
                object fileFormat = WdSaveFormat.wdFormatXMLDocument;
                doc.SaveAs2(ref fileName, ref fileFormat);
                MessageBox.Show("Chữ ký đã được lưu tại " + fileName, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                // Đóng tài liệu và thoát ứng dụng Word
                doc.Close();
                word.Quit();
            }

        }


        private void btnFileBanRoCheck_Click(object sender, EventArgs e)
        {
            txtBanRoCheck.Text = upFile();
            if (this.txtBanRoCheck.Text.Equals(""))
            {
                MessageBox.Show("Chưa có văn bản nào được nhập vào!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void btnFileChuKyCheck_Click(object sender, EventArgs e)
        {
            txtChuKyCheck.Text = upFile();
            if (this.txtChuKyCheck.Text.Equals(""))
            {
                MessageBox.Show("Chưa có văn bản nào được nhập vào!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        //Function decrypt RSA
        public string giaiMaRSA(string[] banMa, BigInteger n, BigInteger b)
        {
            string result = "";

            foreach (string item in banMa)
            {
                string temp = thuatToanBinhPhuongVaNhanCheck(item, n, b);

                BigInteger x = BigInteger.Parse(temp);
                char c;
                try
                {
                    c = (char)x;
                }
                catch (Exception)
                {
                    return result;
                }


                result = result + c.ToString();
            }

            return result;
        }

        public string thuatToanBinhPhuongVaNhanCheck(string ch, BigInteger m, BigInteger b)
        {
            BigInteger c = BigInteger.Parse(ch);

            //convert số mũ qua dạng nhị phân
            string charNhiPhan = toBinary(b, 2);
            char[] mangNhiPhan = charNhiPhan.ToCharArray();
            BigInteger p = 1;
            foreach (char item in mangNhiPhan)
            {
                p = BigInteger.Pow(p, 2);
                p = BigInteger.Remainder(p, m);
                if (item == '1')
                {
                    p = BigInteger.Multiply(p, c);
                    p = BigInteger.Remainder(p, m);
                }
            }

            return p.ToString();
        }
        private void btnCheck_Click(object sender, EventArgs e)
        {
            if (this.txtBanRoCheck.Text.Trim().Equals(""))
            {
                this.txtThongBao.Text = "";
                MessageBox.Show("Chưa có thông điệp để kiểm tra!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (this.txtChuKyCheck.Text.Trim().Equals(""))
            {
                this.txtThongBao.Text = "";
                MessageBox.Show("Chưa có chữ ký để kiểm tra!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (this.txtP.Text.Trim().Equals(""))
            {
                this.txtThongBao.Text = "";
                MessageBox.Show("Vui lòng nhập P!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (this.txtQ.Text.Trim().Equals(""))
            {
                this.txtThongBao.Text = "";
                MessageBox.Show("Vui lòng nhập Q!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.txtThongBao.Text = "";
                return;
            }
            if (this.txtB.Text.Trim().Equals(""))
            {
                this.txtThongBao.Text = "";
                MessageBox.Show("Vui lòng nhập B!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.txtThongBao.Text = "";
                return;
            }
            if (this.txtA.Text.Trim().Equals(""))
            {
                this.txtThongBao.Text = "";
                MessageBox.Show("Vui lòng nhập A!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.txtThongBao.Text = "";
                return;
            }


            BigInteger n = BigInteger.Multiply(BigInteger.Parse(this.txtP.Text.Trim()), BigInteger.Parse(this.txtQ.Text.Trim()));
            BigInteger b = BigInteger.Parse(this.txtB.Text);

            string textBanRoCheck = this.txtBanRoCheck.Text;
            string textChuKyCheck = this.txtChuKyCheck.Text;

            //Băm bản rõ để check
            string hashBanRoCheck = "";
            if (openFileWord == true)
            {
                hashBanRoCheck = hashMD5Word(this.txtBanRoCheck);
            }
            else
            {
                hashBanRoCheck = hashMD5(this.txtBanRoCheck.Text);
            }

            //Đưa bản mã về mảng string để xét
            string[] arrayBanMaCheck = textChuKyCheck.Split('-');

            ///ko cần
            string[] md5Char = new string[32];
            int i = 0;
            foreach (var item in arrayBanMaCheck)
            {
                md5Char[i] = item;
                i++;
            }

            string giaiMa = giaiMaRSA(md5Char, n, b);
            if (openFileWord)
                openFileWord = false;

            if (hashBanRoCheck.Trim().Equals(giaiMa.Trim()) || giaiMa.Equals(hashMD5Word(this.txtBanRoCheck)))
            {
                this.txtThongBao.Text = "Chữ ký chính xác!";
                return;
            }

            else
            {
                this.txtThongBao.Text = "Thông điệp, chữ ký hoặc P,Q đã bị thay đổi. Vui lòng kiểm tra lại!";
                return;
            }

        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            this.txtA.ResetText();
            this.txtB.ResetText();
            this.txtP.ResetText();
            this.txtQ.ResetText();
            this.txtBanRo.ResetText();
            this.txtBanRoCheck.ResetText();
            this.txtChuKy.ResetText();
            this.txtChuKyCheck.ResetText();
            rsa = null;
            txtHashMD5 = "";
            this.txtThongBao.Text = "";
            openFileWord = false;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            this.txtBanRo.ResetText();
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "|*.docx";
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                openFileWord = true;
                MessageBox.Show($"Bạn đã mở thành công file sau: {openFile.FileName}", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                OpenDocxFile(openFile.FileName, this.txtBanRo);

            }
        }

        public void OpenDocxFile(string filePath, RichTextBox richTextBox1)
        {
            Document document = new Document();
            document.LoadFromFile(filePath);

            int indexShape = 0;

            // Đọc văn bản và định dạng kí tự
            foreach (Section section in document.Sections)
            {
                foreach (Paragraph para in section.Paragraphs)
                {
                    string text = "";
                    foreach (DocumentObject obj in para.ChildObjects)
                    {
                        if (obj is TextRange)
                        {

                            richTextBox1.SelectionFont = new Font("Microsoft Sans Serif", 11, FontStyle.Regular);
                            richTextBox1.SelectionCharOffset = 0;
                            richTextBox1.SelectionBackColor = Color.White;
                            richTextBox1.ForeColor = System.Drawing.Color.Black;
                            richTextBox1.BackColor = System.Drawing.Color.White;

                            // Lấy định dạng kí tự của từ
                            TextRange range = obj as TextRange;

                            string fontStyle = range.CharacterFormat.FontName;
                            float fontSize = range.CharacterFormat.FontSize;
                            Color fontColor = range.CharacterFormat.TextColor;
                            Color highlight = range.CharacterFormat.HighlightColor;
                            Color textBackColor = range.CharacterFormat.TextBackgroundColor;

                            richTextBox1.SelectionColor = fontColor;
                            richTextBox1.SelectionBackColor = highlight;
                            richTextBox1.SelectionBackColor = textBackColor;

                            List<FontStyle> styles = new List<FontStyle>();

                            // Xác định định dạng
                            if (range.CharacterFormat.Bold)
                            {
                                styles.Add(FontStyle.Bold);
                            }
                            if (range.CharacterFormat.Italic)
                            {
                                styles.Add(FontStyle.Italic);
                            }
                            if (range.CharacterFormat.UnderlineStyle != UnderlineStyle.None)
                            {
                                styles.Add(FontStyle.Underline);
                            }
                            if (range.CharacterFormat.IsStrikeout)
                            {
                                styles.Add(FontStyle.Strikeout);
                            }

                            FontStyle fontStyle1 = FontStyle.Regular;
                            // Áp dụng định dạng
                            if (styles.Count > 0)
                            {

                                foreach (var style in styles)
                                {
                                    fontStyle1 |= style;
                                }
                                richTextBox1.SelectionFont = new Font(fontStyle, fontSize, fontStyle1);
                            }
                            else
                            {
                                richTextBox1.SelectionFont = new Font(fontStyle, fontSize, FontStyle.Regular);
                            }

                            var superscipt = range.CharacterFormat.SubSuperScript;

                            if (superscipt == SubSuperScript.SuperScript)
                            {
                                richTextBox1.SelectionCharOffset = (int)(fontSize * 0.45);
                                richTextBox1.SelectionFont = new Font(fontStyle, fontSize * 0.6f, fontStyle1);
                            }
                            else if (superscipt == SubSuperScript.SubScript)
                            {
                                richTextBox1.SelectionCharOffset = 0; // Đặt lại vị trí văn bản
                                richTextBox1.SelectionFont = new Font(fontStyle, fontSize * 0.6f, fontStyle1);
                                richTextBox1.SelectionCharOffset = -(int)(fontSize * 0.15); // Di chuyển văn bản xuống

                            }
                            richTextBox1.AppendText(range.Text);


                        }

                        if (obj is DocPicture)
                        {
                            // Cast the element to a DocPicture object
                            DocPicture picture = obj as DocPicture;

                            // Get the image data as a Bitmap object
                            Bitmap image = picture.Image as Bitmap;

                            // Display the image in a RichTextBox control
                            if (image != null)
                            {
                                Clipboard.SetDataObject(image);
                                richTextBox1.Paste();
                            }

                        }

                    }
                    richTextBox1.AppendText(Environment.NewLine);
                }
                Application application = new Application();
                Microsoft.Office.Interop.Word.Document documentInterop = application.Documents.Open(filePath);
                foreach (Shape shape in documentInterop.Shapes)
                {
                    shape.Select();
                    application.Selection.Copy();

                    // Dán đối tượng hình vào RichTextBox
                    richTextBox1.Paste();
                }
                documentInterop.Close();
                application.Quit();
            }
            document.Close();
        }

        private void btnFileWordVBKCheck_Click(object sender, EventArgs e)
        {
            this.txtBanRoCheck.ResetText();
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "|*.docx";
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                openFileWord = true;
                MessageBox.Show($"Bạn đã mở thành công file sau: {openFile.FileName}", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                OpenDocxFile(openFile.FileName, this.txtBanRoCheck);

            }
        }

        private void btnFileWordCKCheck_Click(object sender, EventArgs e)
        {
            this.txtChuKyCheck.ResetText();
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "|*.docx";
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                MessageBox.Show($"Bạn đã mở thành công file sau: {openFile.FileName}", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document document = application.Documents.Open(openFile.FileName);

                string text = document.Content.Text;

                this.txtChuKyCheck.Text = text;

                document.Close();
                application.Quit();

            }
        }



        /*
            private void button2_Click(object sender, EventArgs e)
            {
                this.txtBanRoCheck.ResetText();
                OpenFileDialog openFile = new OpenFileDialog();
                openFile.Filter = "|*.docx";
                if (openFile.ShowDialog() == DialogResult.OK)
                {
                    openFileWord = true;
                    MessageBox.Show($"Bạn đã mở thành công file sau: {openFile.FileName}", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    OpenDocxFile(openFile.FileName, this.txtBanRoCheck);
                }
            }

            private void btnFileWordKTChuKy_Click(object sender, EventArgs e)
            {
                this.txtChuKyCheck.ResetText();
                OpenFileDialog openFile = new OpenFileDialog();
                openFile.Filter = "|*.docx";
                if (openFile.ShowDialog() == DialogResult.OK)
                {
                    MessageBox.Show($"Bạn đã mở thành công file sau: {openFile.FileName}", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
                    Microsoft.Office.Interop.Word.Document document = application.Documents.Open(openFile.FileName);

                    string text = document.Content.Text;

                    this.txtChuKyCheck.Text = text;

                    document.Close();
                    application.Quit();

                }
            }


            public string convertMD5(string input)
            {
                // Chuyển đổi chuỗi thành một mảng byte theo mã ASCII
                byte[] inputBytes = Encoding.ASCII.GetBytes(input);

                // Thêm bit 1 vào cuối chuỗi và thêm các bit 0 để đủ độ dài block
                int initialLength = inputBytes.Length;
                int paddingLength = (448 - (initialLength * 8) % 512 + 512) % 512;
                byte[] paddedInputBytes = new byte[initialLength + paddingLength / 8 + 8];
                Array.Copy(inputBytes, paddedInputBytes, initialLength);
                paddedInputBytes[initialLength] = 0x80;
                Array.Copy(BitConverter.GetBytes((ulong)initialLength * 8), 0, paddedInputBytes, paddedInputBytes.Length - 8, 8);

                // Khởi tạo các biến
                uint a = 0x67452301;
                uint b = 0xefcdab89;
                uint c = 0x98badcfe;
                uint d = 0x10325476;

                // Băm từng block 512 bit
                for (int i = 0; i < paddedInputBytes.Length; i += 64)
                {
                    uint[] buffer = new uint[16];
                    for (int j = 0; j < 16; j++)
                    {
                        buffer[j] = BitConverter.ToUInt32(paddedInputBytes, i + j * 4);
                    }

                    uint aa = a, bb = b, cc = c, dd = d;

                    // Round 1
                    for (int j = 0; j < 16; j++)
                    {
                        int g = j;
                        uint f = (b & c) | ((~b) & d);
                        int[] s = { 7, 12, 17, 22 };
                        int r = s[j % 4];
                        a = b + rotate_left(a + f + buffer[g], r);
                        uint temp = d;
                        d = c;
                        c = b;
                        b = a;
                        a = temp;
                    }

                    // Round 2
                    for (int j = 0; j < 16; j++)
                    {
                        int g = (5 * j + 1) % 16;
                        uint f = (d & b) | ((~d) & c);
                        int[] s = { 5, 9, 14, 20 };
                        int r = s[j % 4];
                        a = b + rotate_left(a + f + buffer[g] + 0x5a827999, r);
                        uint temp = d;
                        d = c;
                        c = b;
                        b = a;
                        a = temp;
                    }

                    // Round 3
                    for (int j = 0; j < 16; j++)
                    {
                        int g = (3 * j + 5) % 16;
                        uint f = b ^ c ^ d;
                        int[] s = { 4, 11, 16, 23 };
                        int r = s[j % 4];
                        a = b + rotate_left(a + f + buffer[g] + 0x6ed9eba1, r);
                        uint temp = d;
                        d = c;
                        c = b;
                        b = a;
                        a = temp;
                    }

                    // Round 4
                    for (int j = 0; j < 16; j++)
                    {
                        int g = (7 * j) % 16;
                        uint f = c ^ (b | (~d));
                        int[] s = { 6, 10, 15, 21 };
                        int r = s[j % 4];
                        a = b + rotate_left(a + f + buffer[g] + 0x8f1bbcdc, r);
                        uint temp = d;
                        d = c;
                        c = b;
                        b = a;
                        a = temp;
                    }

                    // Cộng giá trị mới vào kết quả
                    a += aa;
                    b += bb;
                    c += cc;
                    d += dd;
                }

                // Chuyển đổi kết quả sang chuỗi hexa
                byte[] resultBytes = new byte[] { (byte)a, (byte)(a >> 8), (byte)(a >> 16), (byte)(a >> 24),
                                              (byte)b, (byte)(b >> 8), (byte)(b >> 16), (byte)(b >> 24),
                                              (byte)c, (byte)(c >> 8), (byte)(c >> 16), (byte)(c >> 24),
                                              (byte)d, (byte)(d >> 8), (byte)(d >> 16), (byte)(d >> 24) };
                string result = BitConverter.ToString(resultBytes).Replace("-", "").ToLower();

                // In kết quả
                return result;
            }

            // Hàm xoay bit
            public uint rotate_left(uint value, int shift)
            {
                return ((value << shift) | (value >> (32 - shift)));
            }
        */
    }

}