# WebBrowser
对IWebBrowser2的封装，屏蔽了的脚本异常，支持C++ JS互相调用

```C++
  webBrowser_.ExportToJScript(_T("GetProcessID"), 0, [&](DISPPARAMS*, VARIANT * pResult) {
    *pResult = CComVariant(GetProcessId(NULL));
  });
  webBrowser_.ExportToJScript(_T("ShowMessageBox"), 1, [&](DISPPARAMS* pParams, VARIANT *) {
    ::MessageBox(hWnd, pParams->rgvarg[0].bstrVal, _T(""), MB_OK);
  });


  webBrowser_.CallJScript(_T("TestFunc"), _T("1"), _T("2"));
```

主要代码来自

* https://github.com/ietab/ietab.git
