using System;
using System.IO;
using System.Numerics;
using System.Security.Cryptography;
using System.Text;
using System.Collections;

namespace Nskd.Crypt
{
    public class CryptServiceProvider
    {
        private Byte[] aesKey;
        public CryptServiceProvider() { }
        public Byte[] AesKey
        {
            get
            {
                if (aesKey == null)
                {
                    return null;
                }
                else
                {
                    return (Byte[])aesKey.Clone();
                }
            }
            set
            {
                if (value == null)
                {
                    aesKey = null;
                }
                else
                {
                    aesKey = (Byte[])value.Clone();
                    Array.Resize<Byte>(ref aesKey, 32);
                }
            }
        }
        public Byte[] Decrypt(Byte[] data)
        {
            return Aes.Decrypt(data, aesKey);
        }
        public Byte[] Encrypt(Byte[] data)
        {
            return Aes.Encrypt(data, aesKey);
        }
    }
    public class DiffieHellman
    {
        private BigInteger module;
        private BigInteger base_;
        private BigInteger privateKeyA;
        private BigInteger publicKeyA;
        private BigInteger privateKeyB; // не используется
        private BigInteger publicKeyB;
        private BigInteger commonKey;
        public DiffieHellman() { }
        public DiffieHellman(Int32 moduleSize)
        {
            // хорошо бы, что бы,
            // число module было простым и (module - 1) / 2 тоже простым
            // и base_ ^ (module - 1) = 1 mod(module)
            base_ = Utilities.GetRandom(moduleSize);
            do
            {
                module = Utilities.GetProbablePrime(moduleSize); // вернём если не найдётся ничего лучше
            } while (!(BigInteger.ModPow(base_, (module - 1), module).IsOne));
            do
            {
                privateKeyA = Utilities.GetRandom(moduleSize);
                publicKeyA = BigInteger.ModPow(base_, privateKeyA, module);
            } while (publicKeyA.IsOne);
            privateKeyB = 0;
            publicKeyB = 0;
            commonKey = 0;
        }
        public Byte[] GenerateCommonKey(Byte[] publicKeyB)
        {
            this.publicKeyB = new BigInteger(publicKeyB);
            this.commonKey = BigInteger.ModPow(this.publicKeyB, this.privateKeyA, this.module);
            return commonKey.ToByteArray();
        }
        public void ImportParameters(DiffieHellmanParameters ps)
        {
            module = (ps.Module != null) ? new BigInteger(ps.Module) : 0;
            base_ = (ps.Base != null) ? new BigInteger(ps.Base) : 0;
            privateKeyA = (ps.PrivateKeyA != null) ? new BigInteger(ps.PrivateKeyA) : 0;
            publicKeyA = (ps.PublicKeyA != null) ? new BigInteger(ps.PublicKeyA) : 0;
            privateKeyB = (ps.PrivateKeyB != null) ? new BigInteger(ps.PrivateKeyB) : 0;
            publicKeyB = (ps.PublicKeyB != null) ? new BigInteger(ps.PublicKeyB) : 0;
            commonKey = (ps.CommonKey != null) ? new BigInteger(ps.CommonKey) : 0;
        }
        public DiffieHellmanParameters ExportParameters()
        {
            DiffieHellmanParameters dh = new DiffieHellmanParameters();
            dh.Module = module.ToByteArray();
            dh.Base = base_.ToByteArray();
            dh.PrivateKeyA = privateKeyA.ToByteArray();
            dh.PublicKeyA = publicKeyA.ToByteArray();
            dh.PrivateKeyB = privateKeyB.ToByteArray();
            dh.PublicKeyB = publicKeyB.ToByteArray();
            dh.CommonKey = commonKey.ToByteArray();
            return dh;
        }
    }
    public class DiffieHellmanParameters
    {
        public Byte[] Module;
        public Byte[] Base;
        public Byte[] PrivateKeyA;
        public Byte[] PublicKeyA;
        public Byte[] PrivateKeyB; // не используется
        public Byte[] PublicKeyB;
        public Byte[] CommonKey;
    }
    public static class Utilities
    {
        private static Random random = new Random();
        private static class primes
        {
            // простые числа для проверки других на простоту
            public static Int16[] p500 = {
               2,   3,   5,   7,  11,  13,  17,  19,  23,  29,  31,  37,  41,  43,  47,  53,  59,  61,  67,  71, 
              73,  79,  83,  89,  97, 101, 103, 107, 109, 113, 127, 131, 137, 139, 149, 151, 157, 163, 167, 173, 
             179, 181, 191, 193, 197, 199, 211, 223, 227, 229, 233, 239, 241, 251, 257, 263, 269, 271, 277, 281, 
             283, 293, 307, 311, 313, 317, 331, 337, 347, 349, 353, 359, 367, 373, 379, 383, 389, 397, 401, 409, 
             419, 421, 431, 433, 439, 443, 449, 457, 461, 463, 467, 479, 487, 491, 499, 503, 509, 521, 523, 541, 
             547, 557, 563, 569, 571, 577, 587, 593, 599, 601, 607, 613, 617, 619, 631, 641, 643, 647, 653, 659, 
             661, 673, 677, 683, 691, 701, 709, 719, 727, 733, 739, 743, 751, 757, 761, 769, 773, 787, 797, 809, 
             811, 821, 823, 827, 829, 839, 853, 857, 859, 863, 877, 881, 883, 887, 907, 911, 919, 929, 937, 941, 
             947, 953, 967, 971, 977, 983, 991, 997,1009,1013,1019,1021,1031,1033,1039,1049,1051,1061,1063,1069,
            1087,1091,1093,1097,1103,1109,1117,1123,1129,1151,1153,1163,1171,1181,1187,1193,1201,1213,1217,1223,
            1229,1231,1237,1249,1259,1277,1279,1283,1289,1291,1297,1301,1303,1307,1319,1321,1327,1361,1367,1373,
            1381,1399,1409,1423,1427,1429,1433,1439,1447,1451,1453,1459,1471,1481,1483,1487,1489,1493,1499,1511,
            1523,1531,1543,1549,1553,1559,1567,1571,1579,1583,1597,1601,1607,1609,1613,1619,1621,1627,1637,1657,
            1663,1667,1669,1693,1697,1699,1709,1721,1723,1733,1741,1747,1753,1759,1777,1783,1787,1789,1801,1811,
            1823,1831,1847,1861,1867,1871,1873,1877,1879,1889,1901,1907,1913,1931,1933,1949,1951,1973,1979,1987,
            1993,1997,1999,2003,2011,2017,2027,2029,2039,2053,2063,2069,2081,2083,2087,2089,2099,2111,2113,2129,
            2131,2137,2141,2143,2153,2161,2179,2203,2207,2213,2221,2237,2239,2243,2251,2267,2269,2273,2281,2287,
            2293,2297,2309,2311,2333,2339,2341,2347,2351,2357,2371,2377,2381,2383,2389,2393,2399,2411,2417,2423,
            2437,2441,2447,2459,2467,2473,2477,2503,2521,2531,2539,2543,2549,2551,2557,2579,2591,2593,2609,2617,
            2621,2633,2647,2657,2659,2663,2671,2677,2683,2687,2689,2693,2699,2707,2711,2713,2719,2729,2731,2741,
            2749,2753,2767,2777,2789,2791,2797,2801,2803,2819,2833,2837,2843,2851,2857,2861,2879,2887,2897,2903,
            2909,2917,2927,2939,2953,2957,2963,2969,2971,2999,3001,3011,3019,3023,3037,3041,3049,3061,3067,3079,
            3083,3089,3109,3119,3121,3137,3163,3167,3169,3181,3187,3191,3203,3209,3217,3221,3229,3251,3253,3257,
            3259,3271,3299,3301,3307,3313,3319,3323,3329,3331,3343,3347,3359,3361,3371,3373,3389,3391,3407,3413,
            3433,3449,3457,3461,3463,3467,3469,3491,3499,3511,3517,3527,3529,3533,3539,3541,3547,3557,3559,3571
                };
        }
        public static BigInteger GetRandom(int byteSize)
        {
            Byte[] bytes = new Byte[byteSize];
            random.NextBytes(bytes);
            bytes[byteSize - 1] &= 0x3f; // is positive and < Max
            bytes[0] |= 0x01; // is odd
            return new BigInteger(bytes);
        }
        private static BigInteger getRandomMax(int byteSize)
        {
            Byte[] bytes = new Byte[byteSize];
            random.NextBytes(bytes); // is random
            bytes[byteSize - 1] &= 0x7f; // is positive
            bytes[byteSize - 1] |= 0x40; // is max
            bytes[0] |= 0x03; // is odd, (n-1)/2 is odd too
            return new BigInteger(bytes);
        }
        private static bool isProbablePrime(BigInteger v, int byteSize, int r)
        {
            bool isProbablePrime = !v.IsEven;
            // сначала проверяем по таблице простых чисел
            for (int i = 1; (i < primes.p500.Length) && isProbablePrime; i++)
            {
                isProbablePrime = !((v % primes.p500[i]).IsZero);
            }
            /*
             * потом тест Миллера-Рабина
             * Ввод: m > 2, нечётное натуральное число, которое необходимо проверить на простоту;
             *    r — количество раундов.
             * Вывод: составное, означает, что m является составным числом;
             *    вероятно простое, означает, что m с высокой вероятностью является простым числом.
             * Представить m − 1 в виде (2^s)·t, где t нечётно, можно сделать последовательным делением m - 1 на 2.
             * цикл А: повторить r раз:
             * Выбрать случайное целое число a в отрезке [2, m − 2]
             * x ← a^t mod m
             * если x = 1 или x = m − 1, то перейти на следующую итерацию цикла А
             * цикл B: повторить s − 1 раз
             *   x ← x^2 mod m
             *   если x = 1, то вернуть составное
             *   если x = m − 1, то перейти на следующую итерацию цикла A
             * вернуть составное
             * вернуть вероятно простое
             */
            if (isProbablePrime)
            {
                BigInteger n = v - 1;
                //int s = 1; // для нашего случая
                BigInteger t = (n / 2);
                while (r-- > 0) // A
                {
                    BigInteger a = GetRandom(byteSize);
                    a = a % n;
                    BigInteger x = BigInteger.ModPow(a, t, v);
                    if (x.IsOne || (x == n)) continue;
                    // B
                    isProbablePrime = false;
                    break;
                }
            }
            return isProbablePrime;
        }
        public static BigInteger GetProbablePrime(int byteSize)
        {
            BigInteger p;
            long start = DateTime.Now.Ticks;
            long finish = start;
            //long count = 0;
            do
            {
                p = getRandomMax(byteSize); // вернём если не найдётся ничего лучше
                if (isProbablePrime(p, byteSize, 10)) break;
                finish = DateTime.Now.Ticks;
                //count++;
            }
            while ((finish - start) < 5000000); // 0.5 sec
            return p;
        }
        public static BigInteger ModInverse(BigInteger e, BigInteger module)
        {
            BigInteger _f = module;
            BigInteger _e = e;
            BigInteger _q = _f / _e;
            BigInteger _r = _f % _e;
            BigInteger _y0;
            BigInteger _y1 = 1;
            BigInteger _t = (_y1 * _q) % module;
            BigInteger _y2 = (module - _t);
            do
            {
                // перенос
                _f = _e;
                _e = _r;
                _y0 = _y1;
                _y1 = _y2;
                // расчёт
                _q = _f / _e;
                _r = _f % _e;
                _t = (_y1 * _q) % module;
                if (_y0 == _t) break; //_y2 = new BigInteger(0);
                else if (_y0 > _t) _y2 = _y0 - _t;
                else _y2 = (module - (_t - _y0));
            }
            while (!_r.IsOne);
            return _y2;
        }
    }
    public static class Aes
    {
        public static byte[] Decrypt(byte[] data, byte[] key)
        {
            byte[] decryptedBytes = null;
            using (RijndaelManaged alg = new RijndaelManaged())
            {
                alg.Key = key;
                alg.BlockSize = 128;
                alg.Mode = CipherMode.CBC;
                alg.Padding = PaddingMode.Zeros;

                ICryptoTransform decryptor = alg.CreateDecryptor(alg.Key, alg.IV);
                using (MemoryStream ms = new MemoryStream(data))
                {
                    using (CryptoStream cs = new CryptoStream(ms, decryptor, CryptoStreamMode.Read))
                    {
                        byte[] buff = new byte[data.Length];
                        int count = cs.Read(buff, 0, buff.Length);
                        if (count > 16)
                        {
                            for (int i = 0; i < count - 16; i++) buff[i] = buff[i + 16];
                            decryptedBytes = new byte[count - 16];
                            Array.Copy(buff, decryptedBytes, count - 16);
                        }
                    }
                }
            }
            return decryptedBytes;
        }
        public static byte[] Encrypt(byte[] data, byte[] key)
        {
            byte[] encryptedBytes = null;
            using (RijndaelManaged alg = new RijndaelManaged())
            {
                alg.Key = key;
                alg.BlockSize = 128;
                alg.Mode = CipherMode.CBC;
                alg.Padding = PaddingMode.Zeros;

                ICryptoTransform encryptor = alg.CreateEncryptor(alg.Key, alg.IV);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor, CryptoStreamMode.Write))
                    {
                        //  buff = IV + plainBytes + pad16
                        byte[] buff = new byte[((data.Length / 16) + 2) * 16];
                        data.CopyTo(buff, 16);
                        cs.Write(buff, 0, buff.Length);
                        encryptedBytes = ms.ToArray();
                    }
                }
            }
            return encryptedBytes;
        }
    }
    public class Rsa // PKCS#1 v2.1 RSAES-PKCS1-v1_5
    {
        private static Random rnd = new Random();
        private int k; // length in octets of the RSA modulus n
        private BigInteger p; // first prime factors of the RSA modulus n
        private BigInteger q; // second prime factors of the RSA modulus n
        private BigInteger n; // RSA modulus
        private BigInteger e; // RSA public exponent
        private BigInteger d; // RSA private exponent
        private byte[] pad1(byte[] m)
        {
            byte[] pm = new byte[k];
            int psc = k - 3 - m.Length;
            if (psc < 8) throw new ArgumentException();
            int i = 0;
            pm[i++] = 0;
            pm[i++] = 1;
            while (--psc >= 0) pm[i++] = 0xff;
            pm[i++] = 0;
            foreach (byte b in m) pm[i++] = b;
            return pm;
        }
        private byte[] pad2(byte[] m)
        {
            byte[] pm = new byte[k];
            int psc = k - 3 - m.Length;
            if (psc < 8) throw new ArgumentException();
            int i = 0;
            pm[i++] = 0;
            pm[i++] = 2;
            while (--psc >= 0) pm[i++] = (byte)(rnd.Next(1, 256));
            pm[i++] = 0;
            foreach (byte b in m) pm[i++] = b;
            return pm;
        }
        private byte[] unpad1(byte[] pm)
        {
            int i = 0;
            if (pm[i++] != 0) throw new ArgumentException();
            if (pm[i++] != 1) throw new ArgumentException();
            while ((i < pm.Length) && (pm[i++] != 0)) ;
            if (i < 10) throw new ArgumentException();
            int mLen = pm.Length - i;
            byte[] m = new byte[mLen];
            for (int j = 0; j < mLen; j++) m[j] = pm[i++];
            return m;
        }
        private byte[] unpad2(byte[] pm)
        {
            int i = 0;
            if (pm[i++] != 0) throw new ArgumentException();
            if (pm[i++] != 2) throw new ArgumentException();
            while ((i < pm.Length) && (pm[i++] != 0)) ;
            if (i < 10) throw new ArgumentException();
            int mLen = pm.Length - i;
            if (mLen < 0) throw new ArgumentException();
            byte[] m = new byte[mLen];
            for (int j = 0; j < mLen; j++) m[j] = pm[i++];
            return m;
        }
        public Rsa() { }
        public Rsa(int byteSize)
        {
            k = byteSize;
            int pk = k / 2;
            int qk = k - pk;
            p = Utilities.GetProbablePrime(pk);
            q = Utilities.GetProbablePrime(qk);
            n = p * q;
            BigInteger f = (p - 1) * (q - 1);
            do
            {
                e = Utilities.GetRandom(k) % f; // 1 < e < f
            }
            while (!BigInteger.GreatestCommonDivisor(f, e).IsOne);
            d = Utilities.ModInverse(e, f);
        }
        public void ImportParameters(RsaParameters ps)
        {
            if (ps.Module != null)
            {
                k = ps.Module.Length;
                if (ps.P != null) p = new BigInteger(ps.P);
                if (ps.Q != null) q = new BigInteger(ps.Q);
                if (ps.Module != null) n = new BigInteger(ps.Module);
                if (ps.Exponent != null) e = new BigInteger(ps.Exponent);
                if (ps.D != null) d = new BigInteger(ps.D);
            }
        }
        public RsaParameters ExportParameters()
        {
            RsaParameters ps = new RsaParameters();
            ps.P = p.ToByteArray();
            ps.Q = q.ToByteArray();
            ps.Module = n.ToByteArray();
            ps.Exponent = e.ToByteArray();
            ps.D = d.ToByteArray();
            return ps;
        }
        public byte[] EncryptMessage(byte[] message, bool fOAEP)
        {
            byte[] em = pad2(message);
            Array.Reverse(em);
            BigInteger c = BigInteger.ModPow(new BigInteger(em), e, n);
            Byte[] ciphertext = c.ToByteArray();
            return ciphertext;
        }
        public byte[] EncryptSignature(byte[] message)
        {
            Byte[] em = pad1(message);
            Array.Reverse(em);
            BigInteger c = BigInteger.ModPow(new BigInteger(em), d, n);
            Byte[] ciphertext = c.ToByteArray();
            return ciphertext;
        }
        public byte[] DecryptMessage(Byte[] ciphertext, bool fOAEP)
        {
            BigInteger c = new BigInteger(ciphertext);
            Byte[] em = BigInteger.ModPow(c, d, n).ToByteArray();
            Array.Resize(ref em, k);
            Array.Reverse(em);
            Byte[] message = unpad2(em);
            return message;
        }
        public byte[] DecryptSignature(Byte[] ciphertext)
        {
            BigInteger c = new BigInteger(ciphertext);
            Byte[] em = BigInteger.ModPow(c, e, n).ToByteArray();
            Array.Resize(ref em, k);
            Array.Reverse(em);
            Byte[] message = unpad1(em);
            return message;
        }
    }
    public class RsaParameters
    {
        public Byte[] P;
        public Byte[] Q;
        public Byte[] Module;
        public Byte[] Exponent;
        public Byte[] D;
        public RsaParameters() { }
    }
    public class RsaSignHash
    {
        private static Hashtable codes;
        public enum Algs { md2, md5, sha1, sha256, sha384, sha512 };
        public Algs Alg;
        public Byte[] Value;
        public Byte[] Code;
        public Byte[] Info;
        static RsaSignHash()
        {
            codes = new Hashtable();
            codes.Add(Algs.md2,
                new byte[] { 0x30, 0x20, 0x30, 0x0c, 0x06, 0x08, 0x2a, 0x86, 0x48, 0x86, 0xf7, 0x0d, 0x02, 0x02, 0x05, 0x00, 0x04, 0x10 });
            codes.Add(Algs.md5,
                new byte[] { 0x30, 0x20, 0x30, 0x0c, 0x06, 0x08, 0x2a, 0x86, 0x48, 0x86, 0xf7, 0x0d, 0x02, 0x05, 0x05, 0x00, 0x04, 0x10 });
            codes.Add(Algs.sha1,
                new byte[] { 0x30, 0x21, 0x30, 0x09, 0x06, 0x05, 0x2b, 0x0e, 0x03, 0x02, 0x1a, 0x05, 0x00, 0x04, 0x14 });
            codes.Add(Algs.sha256,
                new byte[] { 0x30, 0x31, 0x30, 0x0d, 0x06, 0x09, 0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x01, 0x05, 0x00, 0x04, 0x20 });
            codes.Add(Algs.sha384,
                new byte[] { 0x30, 0x41, 0x30, 0x0d, 0x06, 0x09, 0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x02, 0x05, 0x00, 0x04, 0x30 });
            codes.Add(Algs.sha512,
                new byte[] { 0x30, 0x51, 0x30, 0x0d, 0x06, 0x09, 0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x03, 0x05, 0x00, 0x04, 0x40 });
        }
        public RsaSignHash() { }
        public RsaSignHash(Algs alg, byte[] value)
        {
            Code = (byte[])RsaSignHash.codes[alg];
            Info = new byte[Code.Length + value.Length];
            Code.CopyTo(Info, 0);
            value.CopyTo(Info, Code.Length);
        }
        public RsaSignHash(byte[] info)
        {
            foreach (DictionaryEntry de in codes)
            {
                bool sr = true;
                int i = 0;
                foreach (byte b in (byte[])de.Value) { if (info[i++] != b) { sr = false; break; } }
                if (sr == true)
                {
                    Alg = (Algs)de.Key;
                    Value = new byte[info.Length - i];
                    int j = 0;
                    while (i < info.Length) Value[j++] = info[i++];
                    break;
                }
            }
        }
    }
}
