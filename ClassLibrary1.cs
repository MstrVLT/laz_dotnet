using System;
using System.Runtime.InteropServices;

namespace ClassLibrary1
{
    [ComVisible(true)]
    [Guid("158C2477-6F7D-4B7A-BC58-DDF7AC098109"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ICalculator
    {
        void Sum(int a, int b, out int c);
    }

    [ComVisible(true)]
    [Guid("CE092C8F-3B54-4988-B227-60BD091D194B"), ClassInterface(ClassInterfaceType.None)]
    public class Calculator : ICalculator
    {
        public void Sum(int a, int b, out int c)
        {
            c = a + b;
        }
    }

}
