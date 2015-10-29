using System.Collections.Generic;

namespace DataImporter.Business.Interface
{
    public interface IParser
    {
        /// <summary>
        /// Reads the Line using the Parser Implementation logic.
        /// </summary>
        /// <returns></returns>
        IEnumerable<string[]> Readline();
    }
}
