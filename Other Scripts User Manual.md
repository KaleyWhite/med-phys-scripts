# Other Scripts User Manual
This is a how-to guide for end users of Kaley's scripts that are _not_ RayStation scripts. For RayStation scripts, see the [RayStation Scripts User Manual](./RayStation Scripts User Manual.md).
## ChromeDefaultBrowser
We recently had an issue with our default browser. Every day when we logged in, we set Chrome as the default browser. But the next morning, the default browser was reset to Edge. It turned out that Group Policy had started resetting everyone's default browser to Edge upon logout. Therefore, in order to use a different default browser, we must manually set it every time we log in. Here is a way to automatically set Chrome as your default browser every time you log in. Like all good IT workarounds, it doesn't require admin permissions.

I got this solution from [Stack Exchange](https://superuser.com/questions/15596/automatically-run-a-script-when-i-log-on-to-windows), [Mark McClelland](https://poetengineer.postach.io/), and [Christoph Kolbicz](https://kolbi.cz/).

## Dependencies
- The browser you want to set as the default. To use the solution as is, this is Chrome.
- Christoph Kolbicz's [`SetDefaultBrowser.exe`](https://kolbi.cz/blog/2017/11/10/setdefaultbrowser-set-the-default-browser-per-user-on-windows-10-and-server-2016-build-1607/?sfw=pass1643728209)

This solution was tested in Windows 10, so it is not guaranteed to work with any other operating system.

## Implementation
The batch file `ChromeDefaultBrowser.bat` runs `SetDefaultBrowser.exe` (using the absolute path) with the `chrome` argument:
```bat
START T:\Physics\KW\med-phys-scripts\ChromeDefaultBrowser\SetDefaultBrowser.exe chrome
```
The file `ChromeDefaultBrowser.xml` is ean exported task from Task Scheduler. Import it into Task Scheduler. The task runs the batch file whenever:
- your computer is unlocked
- you log in
- you connect to a user session 
The following are user specific and should be changed:
- The path to `SetDefaultBrowser.exe` in the batch file. I recommend using an absolute path.
- The path to the batch file in the XML file. The path is inside the `<Command>` tags.
- Your username in the XML file. This is between the `<Author>` and the `<UserId>` tags, where the placeholder is `DOMAIN\USER`.