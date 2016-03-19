using System.Collections.Generic;

namespace QuestionPaperGenerator
{
	class QuestionItem
	{
		public string Id { get; set; }
		public string Question { get; set; }
		public List<string> Options { get; set; }
		public string Answer { get; set; }

		public override string ToString()
		{
			return string.Format("Question: {0}, Options: {1}, Answer: {2}", Question, Options, Answer);
		}
	}
}