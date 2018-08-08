using System;

namespace JsonExcel
{
	public static class JsonExcelException
	{
		private const string LOAD_FILE_MESSAGE = "Error: Could not read file from disk. Original error:";
		private const string FILE_NOT_FOUND_MESSAGE = "Error: File does not exits";
		private const string PARSE_MESSAGE = "Error: Could not parse File";

		public static JsonExcelLoadFileException LoadFileException(Exception innerexception)
		{
			return new JsonExcelLoadFileException(LOAD_FILE_MESSAGE, innerexception);
		}

		public static JsonExcelLoadFileException LoadFileNoFileFoundException()
		{
			return new JsonExcelLoadFileException(FILE_NOT_FOUND_MESSAGE);
		}

		public static JsonExcelParseException ParseBranchException(string message, Exception innerexception)
		{
			return new JsonExcelParseException(message);
		}

		public static JsonExcelParseException ParseFileException()
		{
			return new JsonExcelParseException(PARSE_MESSAGE);
		}
	}

	public class JsonExcelLoadFileException : Exception
	{
		public override string Message => base.Message;

		public JsonExcelLoadFileException(string message) : base(message)
		{ }

		public JsonExcelLoadFileException(string message, Exception innerexception) : base(message, innerexception)
		{
		}
	}

	public class JsonExcelSaveFileException : Exception
	{
	}

	public class JsonExcelParseException : Exception
	{
		public JsonExcelParseException(string message) : base(message)
		{
		}
	}
}