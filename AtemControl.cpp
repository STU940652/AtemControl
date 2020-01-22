// AtemControl.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <iostream>
#include <locale>
#include <codecvt>
#include <string>

#include "BMDSwitcherAPI_h.h"

class CmdArgs {
public:
    std::string ip = std::string("");
    std::string run_macro = std::string("");
    bool list_macros = false;
    bool success = false;

    CmdArgs(int argc, char* argv[]) {
        for (int i = 1; i < argc; i++) {
            if (_strcmpi(argv[i], "--ip") == 0) {
                i++;
                if ((i < argc) && (argv[i][0] != '-')) {
                    ip = argv[i];
                    std::cout << "Set Device IP to: " + ip << std::endl;
                    continue;
                } else {
                    std::cout << "ERROR: --ip missing ip address" << std::endl;
                    i--;
                    return;
                }
            }

            if (_strcmpi(argv[i], "--run_macro") == 0) {
                i++;
                if ((i < argc) && (argv[i][0] != '-')) {
                    run_macro = argv[i];
                    continue;
                }
                else {
                    std::cout << "ERROR: --run_macro missing name" << std::endl;
                    i--;
                    return;
                }
            }

            if (_strcmpi(argv[i], "--list_macros") == 0) {
                list_macros = true;
                continue;
            }

            std::cout << "ERROR: Unknown argument: " << argv[i] << std::endl;
            return;
        }
        success = true;
    }

    void list_args(void) {
        std::cout << "Command Line Argument:" << std::endl;
        std::cout << "           ip: " << ip << std::endl;
        std::cout << "    run_macro: " << run_macro << std::endl;
        std::cout << "  list_macros: " << (list_macros ? "True" : "False") << std::endl;
        std::cout << "      success: " << (success ? "True" : "False") << std::endl;
    }
};

void print_usage(void) {
    std::cout << std::endl;
    std::cout << "Command line arguments" << std::endl;
    std::cout << "----------------------" << std::endl;
    std::cout << "  --ip <address>      Set target device IP address" << std::endl;
    std::cout << "  --run_macro <name>  Run a named macro" << std::endl;
    std::cout << "  --list_macros       List Named macros" << std::endl;
    std::cout << std::endl;

}

int main(int argc, char* argv[], char* envp[])
{
    // Variable Declaration
    IBMDSwitcherDiscovery* mSwitcherDiscovery = NULL;
    IBMDSwitcher* mSwitcher = NULL;
    IBMDSwitcherMixEffectBlock* mMixEffectBlock = NULL;
    HRESULT hr;

    std::cout << std::endl;
    std::cout << "AtemControl" << std::endl;
    std::cout << "-----------" << std::endl;
    std::cout << "For control of ATEM M/E Video Switchers" << std::endl;

    // Parse arguments
    CmdArgs args = CmdArgs(argc, argv);
    //args.list_args();
    if (!args.success || argc == 1) {
        print_usage();
        return -1;
    }

    // Initialize COM and Switcher related members
    if (FAILED(CoInitialize(NULL)))
    {
        std::cout << "CoInitialize failed." << std::endl;
        goto bail;
    }

    hr = CoCreateInstance(CLSID_CBMDSwitcherDiscovery, NULL, CLSCTX_ALL, IID_IBMDSwitcherDiscovery, (void**)&mSwitcherDiscovery);
    if (FAILED(hr))
    {
        std::cout << "ERROR: Could not create Switcher Discovery Instance.\nATEM Switcher Software may not be installed.\n";
        goto bail;
    }

    // Do Stuff
    if (args.ip.size()) {

        BMDSwitcherConnectToFailure			failReason;
        std::wstring ws(args.ip.begin(), args.ip.end());
        BSTR addressBstr = SysAllocStringLen(ws.data(), ws.size());

        // Connect to the switcher
        // Note that ConnectTo() can take several seconds to return, both for success or failure,
        // depending upon hostname resolution and network response times, so it may be best to
        // do this in a separate thread to prevent the main GUI thread blocking.
        hr = mSwitcherDiscovery->ConnectTo(addressBstr, &mSwitcher, &failReason);
        SysFreeString(addressBstr);
        if (!SUCCEEDED(hr))
        {
            switch (failReason)
            {
            case bmdSwitcherConnectToFailureNoResponse:
                std::cout << "No response from Switcher" << std::endl;
                break;
            case bmdSwitcherConnectToFailureIncompatibleFirmware:
                std::cout << "Switcher has incompatible firmware" << std::endl;
                break;
            default:
                std::cout << "Connection failed for unknown reason" << std::endl;
                break;
            }
            goto bail;
        }

        if (args.list_macros) {
            // Get list of macros
            IBMDSwitcherMacroPool* mMacroPool;
            const IID iMacroPool = IID_IBMDSwitcherMacroPool;
            mSwitcher->QueryInterface(IID_IBMDSwitcherMacroPool, (void**)&mMacroPool);
            uint32_t macro_count;
            BOOL macro_valid;
            if (FAILED(mMacroPool->GetMaxCount(&macro_count))) {
                std::cout << "ERROR: Failed to get Macro Count" << std::endl;
                goto bail;
            }
            for (uint32_t i = 0; i < macro_count; i++) {
                if (FAILED(mMacroPool->IsValid(i, &macro_valid))) {
                    std::cout << "ERROR: Could not get macro valid flag" << std::endl;
                    goto bail;
                }
                if (macro_valid) {
                    BSTR macro_name;
                    BSTR macro_description;
                    if (FAILED(mMacroPool->GetName(i, &macro_name))) {
                        std::cout << "ERROR getting macro name" << std::endl;
                        goto bail;
                    }

                    if (FAILED(mMacroPool->GetDescription(i, &macro_description))) {
                        std::cout << "ERROR getting macro description" << std::endl;
                        goto bail;
                    }

                    std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>> converter;
                    std::cout << i << " " << converter.to_bytes(macro_name) << " " << converter.to_bytes(macro_description) << std::endl;

                    SysFreeString(macro_name);
                    SysFreeString(macro_description);
                }
            }
        }

        if (args.run_macro.size()) {
            // Run a macro

            uint32_t macro_index = 0;
            bool macro_found = false;

            IBMDSwitcherMacroPool* mMacroPool;
            mSwitcher->QueryInterface(IID_IBMDSwitcherMacroPool, (void**)&mMacroPool);
            uint32_t macro_count;
            BOOL macro_valid;
            if (FAILED(mMacroPool->GetMaxCount(&macro_count))) {
                std::cout << "ERROR: Failed to get Macro Count" << std::endl;
                mMacroPool->Release();
                goto bail;
            }
            for (uint32_t i = 0; i < macro_count; i++) {
                if (FAILED(mMacroPool->IsValid(i, &macro_valid))) {
                    std::cout << "ERROR: Could not get macro valid flag" << std::endl;
                    mMacroPool->Release();
                    goto bail;
                }
                if (macro_valid) {
                    BSTR macro_name;
                    if (FAILED(mMacroPool->GetName(i, &macro_name))) {
                        std::cout << "ERROR getting macro name" << std::endl;
                        mMacroPool->Release();
                        goto bail;
                    }

                    std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>> converter;

                    if (args.run_macro.compare(converter.to_bytes(macro_name)) == 0) {
                        macro_found = true;
                        macro_index = i;
                    }

                    SysFreeString(macro_name);

                    if (macro_found) break;
                }
            }

            if (macro_found) {
                std::cout << "Running macro " << args.run_macro << std::endl;
                // Run macro
                IBMDSwitcherMacroControl* mMacroControl;
                mSwitcher->QueryInterface(IID_IBMDSwitcherMacroControl, (void**)&mMacroControl);

                if (FAILED(mMacroControl->Run(macro_index))) {
                    std::cout << "ERROR: Failed to run macro " << macro_index << std::endl;
                    mMacroControl->Release();
                    goto bail;
                }
                mMacroControl->Release();
                // Delay for a second to make sure the command went through
                Sleep(1000);
            }
            else {
                std::cout << "ERROR: Could not find macro " << args.run_macro << std::endl;
                goto bail;
            }

            mMacroPool->Release();
        }
    }

    bail:
    // Cleanup to close
    if (mSwitcherDiscovery != NULL)
    {
        mSwitcherDiscovery->Release();
        mSwitcherDiscovery = NULL;
    }

    CoUninitialize();
}
