// Библиотеки для работы с excel

namespace eLIBRARYparsing
{
    public struct Vector2Int
    {
        private int _x;
        private int _y;

        public Vector2Int(int x, int y)
        {
            _x = x;
            _y = y;
        }

        public int X => _x;
        public int Y => _y;
    }
}
