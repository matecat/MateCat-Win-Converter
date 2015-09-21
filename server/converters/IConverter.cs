namespace Translated.MateCAT.LegacyOfficeConverter.Converters
{
    public interface IConverter
    {
        /// <summary>
        /// Tries to convert the source file to the target file. Returns false if
        /// the conversion is not supported, true if the conversion succeded, or
        /// raises Exceptions if a supported conversion didn't end well.
        /// 
        /// Before doing the conversion, this function must check the format of the
        /// source file and the format of the target, and decide if it can handle
        /// the required conversion or not. Try to make this check the fastest
        /// possible.
        /// 
        /// For example, if a WordConverter receives an Excel conversion, this
        /// function must return false.
        /// 
        /// This is because this interface is meant to create a "Plugin" like 
        /// structure, where you can easily add or remove converters. To add a new
        /// converter to this conversion server, you can simply add it to the list
        /// in ConvertersRouter. When ConvertersRouter receives a conversion task,
        /// he tries to delegate it to every IConverter in his list, in the
        /// specified order. This is the reason why converters must decide if they
        /// support the requested conversion and quickly return false if they don't.
        /// The first converter that supports the conversion wins the task.
        /// 
        /// See ConvertersRouter code for more info.
        /// </summary>
        bool Convert(string sourceFilePath, int sourceFormat, string targetFilePath, int targetFormat);
    }
}
