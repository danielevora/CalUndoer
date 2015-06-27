using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CalUndoer
{
	class Program
	{
		static string[] Scopes = { CalendarService.Scope.Calendar };
		static string ApplicationName = "CalUndoer";

		static void Main(string[] args)
		{
			UserCredential credential;

			using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
			{
				string credPath = System.Environment.GetFolderPath(
					System.Environment.SpecialFolder.Personal);
				credPath = Path.Combine(credPath, ".credentials");

				credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
					GoogleClientSecrets.Load(stream).Secrets,
					Scopes,
					"user",
					CancellationToken.None,
					new FileDataStore(credPath, true)).Result;
				Console.WriteLine("Credential file saved to: " + credPath);
			}

			// Create Google Calendar API service.
			var service = new CalendarService(new BaseClientService.Initializer()
			{
				HttpClientInitializer = credential,
				ApplicationName = ApplicationName,
			});

			// Define parameters of request.
			EventsResource.ListRequest request = service.Events.List("primary");
			request.TimeMin = DateTime.Now.AddYears(-1);//DateTime.Now.AddYears(-7);
			request.TimeMax = DateTime.Now;
			request.ShowDeleted = false;
			request.SingleEvents = true;
			request.MaxResults = 1000;//{Default:25, Max:2500}
			request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

			// List events.
			Events events = request.Execute();
			Console.WriteLine(String.Format("Event Count: {0}", events.Items.Count));
			if (events.Items != null && events.Items.Count > 0)
			{
				foreach (var eventItem in events.Items)
				{
					Thread.Sleep(200);
					eventItem.Visibility = "private";
					eventItem.Transparency = "transparent";
					EventsResource.UpdateRequest updateReq = service.Events.Update(eventItem, "primary", eventItem.Id);
					Event updatedEvent = updateReq.Execute();
					Console.WriteLine("Event Updated: " + updatedEvent.Id + " " + updatedEvent.Visibility + " " + updatedEvent.Description);
				}
			}
			else
			{
				Console.WriteLine("No events found.");
			}
			Console.Read();
			Console.Read();
		}
	}
}