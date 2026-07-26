#include "MainWindow.h"
#include <QApplication>
#include <QPixmap>
#include <QSplashScreen>
#include <QStyleFactory>
#include <QThread>
int main(int argc,char**argv){
 QApplication app(argc,argv); app.setApplicationName("Excel Compare Professional"); app.setApplicationVersion("1.4.0"); app.setOrganizationName("AAT-Tech Ltd"); app.setWindowIcon(QIcon(":/branding/assets/app-icon.png")); app.setStyle(QStyleFactory::create("Fusion"));
 QPixmap pix(":/branding/assets/splash.png"); QSplashScreen splash(pix.scaled(960,640,Qt::KeepAspectRatio,Qt::SmoothTransformation)); splash.show(); splash.showMessage("Loading Excel Compare Professional...",Qt::AlignBottom|Qt::AlignHCenter,Qt::white); app.processEvents(); QThread::msleep(650);
 MainWindow w; w.show(); splash.finish(&w); return app.exec(); }
