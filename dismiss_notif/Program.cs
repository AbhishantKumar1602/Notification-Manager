using System;
using System.Threading.Tasks;
using Windows.UI.Notifications;
using Windows.UI.Notifications.Management;

class Program
{
    static async Task Main(string[] args)
    {
        var listener = UserNotificationListener.Current;

        var access = await listener.RequestAccessAsync();
        if (access != UserNotificationListenerAccessStatus.Allowed)
        {
            Console.WriteLine("ACCESS_DENIED");
            return;
        }

        var notifications = await listener.GetNotificationsAsync(NotificationKinds.Toast);

        foreach (var n in notifications)
        {
            // Optional text filter
            if (args.Length > 0)
            {
                var visual = n.Notification.Visual;
                var text = "";

                foreach (var b in visual.Bindings)
                    foreach (var t in b.GetTextElements())
                        text += t.Text + " ";

                if (!text.ToLower().Contains(args[0].ToLower()))
                    continue;
            }

            // ✅ CORRECT way to dismiss
            listener.RemoveNotification(n.Id);

            Console.WriteLine($"DISMISSED:{n.Id}");
        }
    }
}
