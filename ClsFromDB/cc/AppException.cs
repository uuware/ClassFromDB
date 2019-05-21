/*Modify List:
 * 20050725,new
 * 
 */
using System;

	/// <summary>
	/// Common Class
	/// </summary>
namespace cc
{
	class AppException : System.Exception
	{
		private bool bisIgnore = false;
		private object tag = null;

		/// <summary>
		/// Standard constructor
		/// </summary>
		/// <param name="message"></param>
		public AppException() : base()
		{
		}

		/// <summary>
		/// Standard constructor
		/// </summary>
		/// <param name="message"></param>
		public AppException(string message) : base(message)
		{
		}

		/// <summary>
		/// Standard constructor
		/// </summary>
		/// <param name="message"></param>
		public AppException(Exception exp) : base(exp.Message, exp.InnerException)
		{
		}

		/// <summary>
		/// Standard constructor
		/// </summary>
		/// <param name="message">The error message that explains
		/// the reason for the exception</param>
		/// <param name="inner">The exception that caused the
		/// current exception</param>
		public AppException(string message, Exception inner) : base(message, inner)
		{
		}

		/// <summary>
		/// Serialization Constructor
		/// </summary>
		protected AppException(System.Runtime.Serialization.SerializationInfo info, System.Runtime.Serialization.StreamingContext context) : base(info, context)
		{
		}

		/// <summary>
		/// Standard constructor
		/// </summary>
		/// <param name="message"></param>
		public AppException(bool bisIgnore) : base()
		{
			this.bisIgnore = bisIgnore;
		}

		/// <summary>
		/// Standard constructor
		/// </summary>
		/// <param name="message"></param>
		public AppException(string message, bool bisIgnore) : base(message)
		{
			this.bisIgnore = bisIgnore;
		}

		/// <summary>
		/// isIgnore�̐ݒ�Ǝ擾
		/// </summary>
		public bool isIgnore
		{
			get
			{
				return bisIgnore;
			}
			set
			{
				bisIgnore = value;
			}
		}

		/// <summary>
		/// isIgnore�̐ݒ�Ǝ擾
		/// </summary>
		public object Tag
		{
			get
			{
				return tag;
			}
			set
			{
				tag = value;
			}
		}

		/// <summary>
		/// MessageAll
		/// </summary>
		public string MessageAll
		{
			get
			{
				return MessageAll_loop(this);
			}
		}

		/// <summary>
		/// get All Message
		/// </summary>
		protected string MessageAll_loop(Exception exp)
		{
			if(exp.InnerException != null)
			{
				return exp.Message + "\r\n" + MessageAll_loop(exp.InnerException);
			}
			else
			{
				return exp.Message;
			}
		}

	}

}
