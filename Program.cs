using System.Net;
using System.Net.Mail;


namespace TimeSheetRemainder
{
    class TimeSheetRemainder
    {
        public static void Main()
        {
            RemainderMail testClass = new RemainderMail();
            Console.WriteLine("Welcome User Management");
            testClass.sendMail();
        }
    }
}