using ResourcesAutogenerate;

namespace ResxPackage.Dialog
{
    public class CombinedLogger: ILogger
    {
        private readonly ILogger _outputWindowLogger;
        private readonly ILogger _dialogLogger;

        public CombinedLogger(ILogger outputWindowLogger, ILogger dialogLogger)
        {
            _outputWindowLogger = outputWindowLogger;
            _dialogLogger = dialogLogger;
        }

        public void Log(string message)
        {
            _dialogLogger.Log(message);
            _outputWindowLogger.Log(message);
        }
    }
}
