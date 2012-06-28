using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace CDemoLib
{
    [Guid("63625789-9D21-3BC4-AD43-59B86E949282"),
      InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [ComVisible(true)]
    public interface _COMDemo
    {
        int Value1{get; set;}
        String Name{get; set;}
        int PlusFive(int inval);
    }

    [Guid("8995D837-7DBC-3A6A-BB6A-29ECE026E255"),
       ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    public class COMDemo : _COMDemo
    {
        private int value1;
        private String name;

        public int Value1
        {
            get { return value1; }
            set { value1 = value; }
        }

        public String Name {
            get { return name; }
            set { name = value; }
        }

        public int PlusFive(int inval) {
             inval = inval + 5;
             return inval;
        }
     }
}
