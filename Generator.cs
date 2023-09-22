using System.Text;

namespace ExcelCrack
{
    public class Generator
    {
        static char[] characters = {
                        '0','1','2','3','4','5','6','7','8','9','a','b','c','d','e','f','g','h','i','j' ,'k','l','m','n','o','p',
                        'q','r','s','t','u','v','w','x','y','z','A','B','C','D','E','F','G','H','I','J','C','L','M','N','O','P',
                        'Q','R','S','T','U','V','X','Y','Z','!','?',' ','*','-','+'
        };
        private int[] _counter;
        private int _length;
        private object once = new object();

        public Generator(int length)
        {

            _counter = new int[length];
            _length = length;
        }

        public string generateNextPassword()
        {
            lock (once)
            {
                var result = new StringBuilder();
                for (int i = 0; i < _length; i++)
                {
                    var counter = _counter[i];
                    result.Append(characters[counter]);
                }


                IncrementCounter(_length - 1);


                return result.ToString();
            }
        }

        private void IncrementCounter(int index)
        {

            if (_counter[index] >= characters.Length - 1)
            {
                _counter[index] = 0;
                if (index > 0)
                {
                    IncrementCounter(index - 1);
                }
                else
                {
                    _length++;
                    _counter = new int[_length];
                }
            }
            else
            {
                _counter[index]++;
            }
        }
    }
}
