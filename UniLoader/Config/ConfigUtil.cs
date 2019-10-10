namespace UniLoader.Config
{
    /// <summary>
    /// Базовый парсер конфигов. Может иметь производные парсеры, пригодные под конкретные виды конфигов.
    /// </summary>
    public abstract class ConfigUtil
    {
        protected string FileName { get; set; }

        protected ConfigUtil(string fileName) => FileName = fileName;

        public abstract void ReadConfig();
    }
}