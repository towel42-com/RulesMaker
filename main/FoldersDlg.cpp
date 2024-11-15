#include "FoldersDlg.h"
#include "FoldersModel.h"
#include "ui_FoldersDlg.h"

#include <QTimer>

CFoldersDlg::CFoldersDlg( QWidget *parent ) :
    QDialog( parent ),
    fImpl( new Ui::CFoldersDlg )
{
    fImpl->setupUi( this );
    fImpl->folders->reload( false );
    setWindowTitle( QObject::tr( "Select Root Folder" ) );
}

CFoldersDlg::~CFoldersDlg()
{
}

QString CFoldersDlg::currentPath() const
{
    return fImpl->folders->selectedPath();
}

QString CFoldersDlg::fullPath() const
{
    return fImpl->folders->selectedFullPath();
}

std::shared_ptr< Outlook::Folder > CFoldersDlg::selectedFolder()
{
    return fImpl->folders->selectedFolder();
}
