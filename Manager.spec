# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(['C:\\Users\\User\\OneDrive\\Desktop\\Current\\Manager.py'],
             pathex=['C:\\Users\\User\\OneDrive\\Desktop\\Current\\EXE'],
             binaries=[],
             datas=[],
             hiddenimports=['babel.numbers', 'selenium','_posixsubprocess','org.python','urllib.getproxies_environment','urllib.proxy_bypass_environment','urllib.proxy_bypass','urllib.getproxies','urllib.urlencode','urllib.unquote_plus','urllib.quote_plus','urllib.unquote','urllib.quote','urllib.pathname2url','_posixshmem','multiprocessing.set_start_method','multiprocessing.get_start_method','multiprocessing.get_context','multiprocessing.TimeoutError','_scproxy','termios','java.lang','multiprocessing.BufferTooShort','multiprocessing.AuthenticationError','asyncio.DefaultEventLoopPolicy','vms_lib ','java','_winreg','readline','org','grp','pwd','posix','resource','pyimod03_importers','_uuid','netbios','win32wnet','__builtin__','ordereddict','_manylinux','urlparse','pkg_resources.extern.pyparsing','win32com.shell','com.sun','com','win32api','win32com','pkg_resources.extern.packaging','pkg_resources.extern.appdirs','sets','UserDict','cdecimal','cPickle','copy_reg ','StringIO','cStringIO','Cookie','cookielib',
	     'urllib2','simplejson','backports.ssl_match_hostname','brotli','Queue','urllib3',"'urllib3.packages.six.moves.urllib'.parse",urllib3.packages.six.moves,'socks','_dummy_threading','typing.io','cryptography','OpenSSL.crypto','cryptography.x509','cryptography.hazmat','OpenSSL','xero_python','xero_python.api_client','xero_python.accounting','thread6','selenium.common','selenium.webdriver','gssapi','dns.exception','dns','olefile','PySide2.QtCore','PySide2','PyQt5.QtCore','cffi','Tkinter','tkFont','ttk'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='Manager',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True )
