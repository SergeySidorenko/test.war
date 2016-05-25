package com.prospectconverter.sync;

import java.net.URI;
import java.util.*;

import microsoft.exchange.webservices.data.autodiscover.AutodiscoverService;
import microsoft.exchange.webservices.data.autodiscover.enumeration.UserSettingName;
import microsoft.exchange.webservices.data.autodiscover.response.GetUserSettingsResponse;
import microsoft.exchange.webservices.data.core.ExchangeServerInfo;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.*;
import microsoft.exchange.webservices.data.core.enumeration.search.LogicalOperator;
import microsoft.exchange.webservices.data.core.enumeration.search.SortDirection;
import microsoft.exchange.webservices.data.core.enumeration.service.ConflictResolutionMode;
import microsoft.exchange.webservices.data.core.enumeration.service.DeleteMode;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.item.Contact;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.*;
import microsoft.exchange.webservices.data.property.definition.ExtendedPropertyDefinition;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

public class TestEWS {
	private static UUID rfPropertySetId = UUID.fromString("03658372-9F96-47b2-A703-B504ED14A220");

	public static void main(String[] args) throws Exception {
		try(ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2)) {
			ExchangeCredentials credentials = new WebCredentials("ms2013@s624787081.onlinehome.us", "A3214df2b-b81f-40ee");
			service.setCredentials(credentials);
			service.setUrl(new URI(autodiscoverUrl()));
//			service.autodiscoverUrl("ms2013@s624787081.onlinehome.us");

/*			Contact contact = new Contact(service);
			contact.setGivenName("Dima");
			contact.setMiddleName ("Mi");
			contact.setSurname("Medvedeff");
			contact.setSubject("Contact Details");

// Specify the company name.
			contact.setCompanyName("Kremlin");
			PhysicalAddressEntry paEntry1 = new PhysicalAddressEntry();
			paEntry1.setStreet("Red Square");
			paEntry1.setCity("Moscow");
			paEntry1.setState("RU");
			paEntry1.setPostalCode("120075");
			paEntry1.setCountryOrRegion("RUSSIA");
			contact.getPhysicalAddresses().setPhysicalAddress(PhysicalAddressKey.Home, paEntry1);
			contact.save();
*/
			Contact co = null;//Contact.bind(service, new ItemId("AAMkADYyYjU2ODIxLTRkZGItNDYzNC1iMDA2LTY0ZDIzMGIyNmM0MQBGAAAAAAAjIl/WDBOxQ7XHaNmftic2BwD12QKU33i3RqgGMQ+xmRQiAAAAAAEOAAD12QKU33i3RqgGMQ+xmRQiAAAEIzlXAAA="));
//			if(co != null)
//				co.delete(DeleteMode.HardDelete);

			try {
				co = Contact.bind(service, new ItemId("AAMkADYyYjU2ODIxLTRkZGItNDYzNC1iMDA2LTY0ZDIzMGIyNmM0MQBGAAAAAAAjIl/WDBOxQ7XHaNmftic2BwD12QKU33i3RqgGMQ+xmRQiAAAAAAEOAAD12QKU33i3RqgGMQ+xmRQiAAAEIzzyAAA="));
			} catch (Exception e) {
				e.printStackTrace();
			}
			if(co != null) {
				co.setMiddleName("-MI+");
/*				co.getPhysicalAddresses().getPhysicalAddress(PhysicalAddressKey.Home).setCity("MI-N-SK");
				Calendar c= GregorianCalendar.getInstance();
				c.set(Calendar.YEAR, 1990);
				co.setBirthday(c.getTime());
				c.set(Calendar.YEAR, 2011);
				co.setWeddingAnniversary(c.getTime());
				EmailAddress ea = new EmailAddress("jdoe@mm.cc");
				co.getEmailAddresses().setEmailAddress(EmailAddressKey.EmailAddress1, ea);
				co.getPhoneNumbers().setPhoneNumber(PhoneNumberKey.HomePhone, "3222-223-3212");
				if(!co.getChildren().contains("2Vova"))
					co.getChildren().add("1Vova");
				co.setJobTitle("Assistant");
				co.setManager("J.F.K.1");
				co.setOfficeLocation("Farg1o+");
				co.setSpouseName("Olga");

				ExtendedPropertyDefinition sourcePropertyDefinition = new ExtendedPropertyDefinition(
						rfPropertySetId, "Source", MapiPropertyType.String);
				co.setExtendedProperty(sourcePropertyDefinition, "RF-defined Property");*/
				//co.update(ConflictResolutionMode.AlwaysOverwrite);
			}
			ExchangeServerInfo info = service.getServerInfo();
			System.out.println(info);
			ItemView view = new ItemView(25);

			view.getOrderBy().add(ItemSchema.LastModifiedTime, SortDirection.Ascending);
			view.getOrderBy().add(ItemSchema.DateTimeCreated, SortDirection.Ascending);
			Calendar from = GregorianCalendar.getInstance();
			from.add(Calendar.MINUTE, -1220);
			from = null;
			FindItemsResults<Item> findResults = null;
			List<Item> itemms = new ArrayList<microsoft.exchange.webservices.data.core.service.item.Item>();
			do {

				findResults = from == null ? service.findItems(WellKnownFolderName.Contacts, view):
						service.findItems(WellKnownFolderName.Contacts,
								new SearchFilter.SearchFilterCollection(
										LogicalOperator.Or,
										new SearchFilter.IsGreaterThan(ItemSchema.LastModifiedTime, from.getTime()),
										new SearchFilter.IsGreaterThan(ItemSchema.DateTimeCreated, from.getTime())),
								view);
				view.setOffset(view.getOffset() + 25);
				itemms.addAll(findResults.getItems());
			} while (findResults.isMoreAvailable());

			System.out.println("?? - " + findResults.getItems().size() + ": " + itemms.size());
			for(Item item : itemms) {
				if(item instanceof microsoft.exchange.webservices.data.core.service.item.Contact) {

					System.out.println(".....................................................................");
					System.out.println("itemID=" + item.getId());
					microsoft.exchange.webservices.data.core.service.item.Contact c = (microsoft.exchange.webservices.data.core.service.item.Contact) item;
//					if(!"AAMkADYyYjU2ODIxLTRkZGItNDYzNC1iMDA2LTY0ZDIzMGIyNmM0MQBGAAAAAAAjIl/WDBOxQ7XHaNmftic2BwD12QKU33i3RqgGMQ+xmRQiAAAAAAEOAAD12QKU33i3RqgGMQ+xmRQiAAAEIzzyAAA=".equals(item.getId().getUniqueId()))
//						c.delete(DeleteMode.HardDelete);
					System.out.println("created=" + c.getDateTimeCreated());
					System.out.println("updated1=" + c.getLastModifiedTime());
					System.out.println("getDisplayName=" + c.getDisplayName());
/*					if(c.getGivenName() != null) {
						c.setGivenName("!" + c.getGivenName());
//						c.update(ConflictResolutionMode.AlwaysOverwrite);
					}

					System.out.println("getGivenName=" + c.getGivenName());
					System.out.println("getMiddleName=" + c.getMiddleName());
					System.out.println("getCompanyName=" + c.getCompanyName());
					System.out.println("getEmailAddresses=" + c.getEmailAddresses());//3 email addresses
					if (c.getEmailAddresses() != null && c.getEmailAddresses().contains (EmailAddressKey.EmailAddress1))
						System.out.println("Email Address=" + c.getEmailAddresses().getEmailAddress(EmailAddressKey.EmailAddress1).getAddress());
					System.out.println("getPhysicalAddresses=" + c.getPhysicalAddresses());
					if (c.getPhysicalAddresses() != null && c.getPhysicalAddresses().getPhysicalAddress(PhysicalAddressKey.Home) != null)
						System.out.println("address City=" + c.getPhysicalAddresses().getPhysicalAddress(PhysicalAddressKey.Home).getCity());
					System.out.println("getPhoneNumbers=" + c.getPhoneNumbers());
					if (c.getPhoneNumbers() != null && c.getPhoneNumbers().getPhoneNumber(PhoneNumberKey.HomePhone) != null)
						System.out.println("Home Phone=" + c.getPhoneNumbers().getPhoneNumber(PhoneNumberKey.HomePhone));
					try {
						System.out.println("getBirthday=" + c.getBirthday());
					} catch (ServiceLocalException e) {
						//e.printStackTrace();
					}
					System.out.println("getBusinessHomePage=" + c.getBusinessHomePage());
					System.out.println("getChildren=" + c.getChildren());
					if (c.getChildren() != null) {
						for (String cld : c.getChildren())
							System.out.println("child=" + cld);
					}
					System.out.println("getOfficeLocation=" + c.getOfficeLocation());
					System.out.println("getManager=" + c.getManager());
					System.out.println("getJobTitle=" + c.getJobTitle());
					System.out.println("getSpouseName=" + c.getSpouseName());
					System.out.println("getSurname=" + c.getSurname());
					try {
						System.out.println("getWeddingAnniversary=" + c.getWeddingAnniversary());
					} catch (ServiceLocalException e) {
						//e.printStackTrace();
					}
					System.out.println("getNotes=" + c.getNotes());
					ExtendedPropertyDefinition sourcePD = new ExtendedPropertyDefinition(
							rfPropertySetId, "Source", MapiPropertyType.String);
//					OutParam<String> src = null;
					try {
						for(ExtendedProperty ep : c.getExtendedProperties())
							if(ep.getPropertyDefinition().getPropertySetId().equals(rfPropertySetId))
                            	System.out.println("EP=" + ep.getValue());
					} catch (Exception e) {
						e.printStackTrace();
					}
					try {
						System.out.println("getContactSource=" + c.getContactSource());//??
					} catch (ServiceLocalException e) {
						//e.printStackTrace();
					}*/
				}
			}
		}
	}

	@SuppressWarnings("incomplete-switch")
	public static String autodiscoverUrl() throws Exception {
		ExchangeCredentials credentials = new WebCredentials("ms2013@s624787081.onlinehome.us", "A3214df2b-b81f-40ee");
		try(AutodiscoverService auto = new AutodiscoverService(new URI("https://exchange.1and1.com/autodiscover/autodiscover.svc"))) {
			auto.setCredentials(credentials);
			GetUserSettingsResponse response = auto.getUserSettings("ms2013@s624787081.onlinehome.us", UserSettingName.ExternalEwsUrl);
			switch (response.getErrorCode()) {
				case NoError:
					return response.getSettings().get(UserSettingName.ExternalEwsUrl).toString();
			}
		}
		return null;
	}

}