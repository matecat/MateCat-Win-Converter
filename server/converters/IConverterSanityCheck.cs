namespace LegacyOfficeConverter
{
    interface IConverterSanityCheck
    {
        /// <summary>
        /// Returns true if the converter is working properly.
        /// The implementation of this method should be the best
        /// balance between speed and robustness. This method
        /// will be called before every conversion so, again, 
        /// make it fast.
        /// </summary
        bool isWorking();
    }
}
