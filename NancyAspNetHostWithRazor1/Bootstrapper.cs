using Nancy;
using Nancy.Conventions;

namespace NancyAspNetHostWithRazor1
{
	public class Bootstrapper : DefaultNancyBootstrapper
	{
		// The bootstrapper enables you to reconfigure the composition of the framework,
		// by overriding the various methods and properties.
		// For more information https://github.com/NancyFx/Nancy/wiki/Bootstrapper
		protected override void ConfigureConventions(NancyConventions nancyConventions)
		{
			nancyConventions.StaticContentsConventions.Add(StaticContentConventionBuilder.AddDirectory("Storage",@"/Storage/Reports"));
			base.ConfigureConventions(nancyConventions);
		}
	}
}