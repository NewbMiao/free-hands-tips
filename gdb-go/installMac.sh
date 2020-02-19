#!/usr/bin/env bash
workspace=$(cd $(dirname $0) && pwd -P)

{
    cd $workspace
    gdb --version
    if [ "$?" = "1" ]; then
        brew install gdb
    fi
    # Start Keychain Access application (/Applications/Utilities/Keychain Access.app)
    # Open the menu item /Keychain Access/Certificate Assistant/Create a Certificate...
    # Choose a name (gdb-cert in the example), set Identity Type to Self Signed Root, set Certificate Type to Code Signing and select the Let me override defaults. Click several times on Continue until you get to the Specify a Location For The Certificate screen, then set Keychain to System.

    # 💡 If you cannot store the certificate in the System keychain: create it in the login keychain instead, then export it. You can then import it into the System keychain.
    # Finally, quit the Keychain Access application to refresh the certificate store.

    security find-certificate -p -c gdb-cert | openssl x509 -checkend 0
    if [ "$?" = "1" ]; then
        echo "Require add gdb-cert to System keychain, pls follow: https://sourceware.org/gdb/wiki/PermissionsDarwin"
        exit
    fi

    codesign --entitlements gdb-entitlement.xml -fs gdb-cert $(which gdb)
    codesign -vv $(which gdb)
    sudo killall taskgated
}
