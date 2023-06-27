using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Numerics;

namespace WinFormsApp7
{
    class RSA
    {
        private BigInteger a, b, n, dn;
        public BigInteger p, q;

        public BigInteger A { get => a; set => a = value; }
        public BigInteger B { get => b; set => b = value; }
        public BigInteger N { get => n; set => n = value; }
        public BigInteger Dn { get => dn; set => dn = value; }

        public RSA(BigInteger p, BigInteger q)
        {
            this.p = p;
            this.q = q;
            this.N = BigInteger.Multiply(p, q);
            this.Dn = BigInteger.Multiply(p - 1, q - 1);
            this.B = randomNumber(Dn);
            this.A = inverseNum(B, Dn);
        }

        public BigInteger randomNumber(BigInteger dn)
        {
            Int64 numMax = 0;

            if (Convert.ToBoolean(dn.CompareTo(65535)))
            {
                numMax = 65525;
            }
            else
            {
                numMax = (Int64)dn;
            }
            Random rd = new Random();
            BigInteger num = rd.NextInt64(2, numMax);

            int gcd = (int)BigInteger.GreatestCommonDivisor(num, dn);

            while (gcd != 1)
            {
                num = rd.NextInt64(2, numMax);
                gcd = (int)BigInteger.GreatestCommonDivisor(num, dn);
            }
            return num;
        }

        public BigInteger inverseNum(BigInteger a, BigInteger m)
        {
            try
            {
                BigInteger r, q, y = 0, y0 = 0, y1 = 1, tga = a, tgm = m;

                while (a > 0)
                {
                    r = m % a;
                    if (r == 0) break;
                    q = m / a;
                    y = y0 - q * y1;
                    m = a;
                    a = r;
                    y0 = y1;
                    y1 = y;
                }
                if (a > 1)
                {
                    return -1;
                }
                if (y > 0)
                    return y;
                else
                    return y + tgm;
            }
            catch (Exception)
            {
                MessageBox.Show("Xin lỗi, Code đang lỗi, vui lòng thử lại!", "Thông báo");
                // this.A = 0;
                return 0;
            }



        }
    }
}
