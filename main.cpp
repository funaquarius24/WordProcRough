#include <QCoreApplication>
#include <QFile>
#include <QJsonDocument>
#include <QJsonObject>
#include <QNetworkAccessManager>
#include <QNetworkReply>

#include <Windows.h>
#include <iostream>
#include <conio.h>
#include <assert.h>
#include <Combaseapi.h>

#include <iostream>
#include <assert.h>

#include "OLE/OLEMethod.h"

int main(int argc, char *argv[]) {
  QCoreApplication a(argc, argv);

  QNetworkAccessManager manager;

  QObject::connect(&manager, &QNetworkAccessManager::finished,
                   [](QNetworkReply *reply) {
                     qDebug() << reply->readAll();
                     delete reply;
                     QCoreApplication::quit();
                   });
  QString base_url("http://127.0.0.1:5000/");
  QString login_url(base_url + "login");
  QNetworkRequest request;
  request.setRawHeader("Content-Type", "application/fhir+json");
  QFile file("themostsimplepatientJSON.json");
  if (!file.open(QIODevice::ReadOnly))
    return -1;

  printf("OK");

  QJsonDocument doc = QJsonDocument::fromJson(file.readAll());
  QJsonObject obj = doc.object();
  obj["id"] = "4705560";
  doc.setObject(obj);

  QJsonDocument login_doc;
  QJsonObject login_info;
  login_info["email"] = "huawei@huawei.com";
  login_info["pwd"] = "123456789";
  login_doc.setObject(login_info);

  request.setUrl(login_url);
  manager.post(request, login_doc.toJson());
  // manager.post(request, &file);

  std::cout << "Reached!!!" << std::endl;

//  CoInitialize(NULL);
//  CLSID clsid;
//  HRESULT hr = CLSIDFromProgID(L"Word.Application", &clsid);
//  // "Excel.Application" for MSExcel

//  IDispatch *pWApp;
//  if(SUCCEEDED(hr))
//  {
//      hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER,
//                            IID_IDispatch, (void **)&pWApp);
//  }

  // Translate server ProgID into a CLSID. ClsidFromProgID
        // gets this information from the registry.
        CLSID clsid;
        CLSIDFromProgID(L"Excel.Application", &clsid);

        // Get an interface to the running instance, if any..
        IUnknown *pUnk;
        HRESULT hr = GetActiveObject(clsid, NULL, (IUnknown**)&pUnk);

        assert(!FAILED(hr));

        // Get IDispatch interface for Automation...
        IDispatch *pDisp;
        hr = pUnk->QueryInterface(IID_IDispatch, (void **)&pDisp);
        assert(!FAILED(hr));

        // Release the no-longer-needed IUnknown...
        pUnk->Release();


  return a.exec();
}
