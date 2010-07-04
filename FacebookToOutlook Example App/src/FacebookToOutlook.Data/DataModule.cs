using System;
using System.Linq;
using Autofac;
using AutoMapper;
using Facebook.Schema;
using FacebookToOutlook.Core;
using FacebookToOutlook.Data.Adapters;

namespace FacebookToOutlook.Data
{
    public class DataModule : Module
    {
        protected override void Load(ContainerBuilder builder)
        {
            builder.RegisterType<OutlookDispatchingRepository>().As<IOutlookRepository>().SingleInstance();
            builder.RegisterType<FacebookRepository>().As<IFacebookRepository>().SingleInstance();

            CreateMappings();
        }

        private static void CreateMappings()
        {
            Mapper.CreateMap<facebookevent, FacebookEvent>()
                .ForMember(e => e.EventId, e => e.MapFrom(fbe => fbe.eid))
                .ForMember(e => e.StartTime, e => e.MapFrom(fbe => ConvertFromFacebookTime(fbe.start_time)))
                .ForMember(e => e.EndTime, e => e.MapFrom(fbe => ConvertFromFacebookTime(fbe.end_time)))
                .ForMember(e => e.EventType, e => e.MapFrom(fbe => fbe.event_type))
                .ForMember(e => e.EventSubType, e => e.MapFrom(fbe => fbe.event_subtype))
                .ForMember(e => e.Host, e => e.MapFrom(fbe => fbe.host))
                .ForMember(e => e.LastModified, e => e.MapFrom(fbe => ConvertFromFacebookTime(fbe.update_time)));
            Mapper.CreateMap<user, FacebookUser>()
                .ForMember(u => u.Name, u => u.MapFrom(usr => usr.name))
                .ForMember(u => u.Company, u => u.MapFrom(usr =>
                                                              {
                                                                  if (usr.work_history == null) return null;
                                                                  var lastWork = usr.work_history.work_info.Last();
                                                                  if (lastWork == null) return null;
                                                                  return lastWork.company_name;
                                                              }))
                .ForMember(u => u.Birthday, u => u.MapFrom(usr =>
                {
                    if (usr.birthday == null)
                        return (DateTime?)null;
                    DateTime val;
                    return DateTime.TryParse(usr.birthday, out val) ? val : (DateTime?) null;
                }))
                .ForMember(u => u.UserId, u => u.MapFrom(usr => usr.uid))
                .ForMember(u => u.PictureUri, u => u.MapFrom(usr => string.IsNullOrEmpty(usr.pic_square) ? null : new Uri(usr.pic_square)));

            Mapper.CreateMap<FacebookEventAdapter, OutlookFacebookEvent>();
            Mapper.CreateMap<IFacebookEvent, FacebookEventAdapter>().ForMember(dest => dest.IsFacebookEvent, opt => opt.Ignore());
            Mapper.CreateMap<FacebookUserAdapter, OutlookFacebookUser>();
            Mapper.CreateMap<IFacebookUser, FacebookUserAdapter>();
        }

        private static DateTime ConvertFromFacebookTime(long facebookTime)
        {
            var facebookTimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");

            return TimeZoneInfo.ConvertTime(new DateTime(1970, 1, 1).AddSeconds(facebookTime),
                TimeZoneInfo.Utc,
                facebookTimeZoneInfo);
        }
    }
}
