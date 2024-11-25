#ifndef WRAPPERS_H
#define WRAPPERS_H

#include <QString>
#include <memory>
class QVariant;
class QWidget;
struct IDispatch;

namespace Outlook
{
    class Application;
    class _Application;
    class NameSpace;
    class Account;
    class _Account;
    class Folder;
    class MAPIFolder;
    class MailItem;
    class _MailItem;
    class AddressEntry;
    class AddressEntries;
    class Recipient;
    class Rules;
    class _Rules;
    class Rule;
    class _Rule;
    class Recipients;
    class AddressList;
    class Items;
    class _Items;

    enum class OlImportance;
    enum class OlItemType;
    enum class OlMailRecipientType;
    enum class OlObjectClass;
    enum class OlRuleConditionType;
    enum class OlSensitivity;
    enum class OlMarkInterval;
}

namespace NWrappers
{
    std::shared_ptr< Outlook::Application > getApplication();

    std::shared_ptr< Outlook::Account > getAccount( Outlook::_Account *item );;

    std::shared_ptr< Outlook::MailItem > getMailItem( IDispatch *item );

    std::shared_ptr< Outlook::Rules > getRules( Outlook::Rules *item );

    std::shared_ptr< Outlook::Rule > getRule( Outlook::_Rule *item );

    std::shared_ptr< Outlook::Folder > getFolder( Outlook::Folder *item );
    std::shared_ptr< Outlook::Folder > getFolder( Outlook::MAPIFolder *item );

    std::shared_ptr< Outlook::Items > getItems( Outlook::_Items *item );
}

#endif