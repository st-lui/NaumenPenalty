using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper.FluentColumnMapping;
using NaumenPenalty.Model;

namespace NaumenPenalty.DataManage
{
    class Mappings
    {
        public static void RegisterDapperMappings()
        {
            var mappings = new ColumnMappingCollection();
            mappings.RegisterType<ServiceCall>()
                .MapProperty(x => x.Id).ToColumn("id")
                .MapProperty(x => x.Removed).ToColumn("removed")
                .MapProperty(x => x.AgreementId).ToColumn("agreement_id")
                .MapProperty(x => x.RegistrationDate).ToColumn("registration_date")
                .MapProperty(x => x.DeadLineTime).ToColumn("deadlinetime")
                .MapProperty(x => x.MassProblem).ToColumn("mass_problem")
                .MapProperty(x => x.Number).ToColumn("number_")
                .MapProperty(x => x.ResolutionTime).ToColumn("resolutiontime")
                .MapProperty(x => x.State).ToColumn("state")
                .MapProperty(x => x.StateStartTime).ToColumn("statestarttime")
                .MapProperty(x => x.ClientouId).ToColumn("clientou_id")
                .MapProperty(x => x.TimeZoneId).ToColumn("timezone_id");
            mappings.RegisterType<Agreement>()
                .MapProperty(x => x.Id).ToColumn("id")
                .MapProperty(x => x.Title).ToColumn("title");
            mappings.RegisterType<Ou>()
                .MapProperty(x => x.Id).ToColumn("id")
                .MapProperty(x => x.LocationId).ToColumn("location");
            mappings.RegisterType<Location>()
                .MapProperty(x => x.Id).ToColumn("id")
                .MapProperty(x => x.PostCode).ToColumn("postoffice$postcode")
                .MapProperty(x => x.Type).ToColumn("postoffice$typecode")
                .MapProperty(x => x.PostalClass).ToColumn("postoffice$postalclass")
                .MapProperty(x => x.Title).ToColumn("title")
				.MapProperty(x=>x.ShortName).ToColumn("postoffice$shortname")
				.MapProperty(x=>x.Parent).ToColumn("parent");
            mappings.RegisterType<NaumenTimeZone>()
                .MapProperty(x => x.Id).ToColumn("id")
                .MapProperty(x => x.Code).ToColumn("code");
            mappings.RegisterWithDapper();
        }
    }
}
/*
				ServiceCall.id
				,ServiceCall.removed
				,ServiceCall.agreement_id
				,ServiceCall.registration_date
				,ServiceCall.deadlinetime
				,ServiceCall.mass_problem
				,ServiceCall.number_
				,ServiceCall.resolutionTime
				,ServiceCall.state
				,ServiceCall.clientou_id
				,ServiceCall.timezone_id
				,Agreement.id
				,Agreement.title
				,Ou.id
				,Ou.location
				,Location.id
				,Location.postoffice$postcode
				,Location.title
*/
