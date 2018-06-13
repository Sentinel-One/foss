//
//  main.m
//  are-you-really-signed
//
//  Created by Matan Mates on 13/06/2018.
//  Copyright © 2018 SentinelOne. All rights reserved.
//

#import <Foundation/Foundation.h>
#include <Security/SecStaticCode.h>

/*
 * Insecure Codesigning Validation as reported originally by Okta
 * https://www.okta.com/security-blog/2018/06/issues-around-third-party-apple-code-signing-checks/
 * Validating with SecStaticCodeCheckValidityWithErrors - kSecCSStrictValidate | kSecCSCheckAllArchitectures
 */

const NSString* ANCHOR_APPLE = @"anchor apple generic";
const NSString* ANCHOR_TRUSTED = @"anchor trusted";

const uint32 ERR_VALID_SIGNATURE      = 0;
const uint32 ERR_SIGNING_EVASION      = 1;
const uint32 ERR_FILE_NOT_SIGNED      = 2;
const uint32 ERR_FILE_NOT_FOUND_USAGE = 3;

const char* VERIFIED_STRN = "\033[92mVERIFIED\033[0m";
const char* FAILED_STRN   = "\033[91mFAILED\033[0m";

bool insecure_validation(NSURL* test_file_url) {
    int rcode = 0;
    SecStaticCodeRef code_ref = NULL;
    
    if (errSecSuccess != SecStaticCodeCreateWithPath((__bridge CFURLRef)test_file_url, kSecCSDefaultFlags, &code_ref)) {
        printf("ERROR: Failed to create codesign object\n");
        return false;
    }
    rcode = SecStaticCodeCheckValidity(code_ref,  kSecCSStrictValidate | kSecCSCheckAllArchitectures, NULL);
    
    CFRelease(code_ref);
    
    return (rcode == errSecSuccess);
}

bool secure_validation(NSURL* test_file_url, bool anchor_trusted) {
    int rcode = 0;
    SecStaticCodeRef code_ref = NULL;
    SecRequirementRef anchor_requirement = NULL;
    
    if (errSecSuccess != SecRequirementCreateWithString(anchor_trusted ? (__bridge CFStringRef)ANCHOR_TRUSTED : (__bridge CFStringRef)ANCHOR_APPLE, kSecCSDefaultFlags, &anchor_requirement)) {
        printf("ERROR: Failed to create anchor requirement\n");
        return false;
    }
    if (errSecSuccess != SecStaticCodeCreateWithPath((__bridge CFURLRef)test_file_url, kSecCSDefaultFlags, &code_ref)) {
        printf("ERROR: Failed to create codesign object\n");
        CFRelease(anchor_requirement);
        return false;
    }
    rcode = SecStaticCodeCheckValidity(code_ref,  kSecCSDefaultFlags | kSecCSCheckNestedCode | kSecCSCheckAllArchitectures | kSecCSEnforceRevocationChecks, anchor_requirement);
    
    CFRelease(code_ref);
    CFRelease(anchor_requirement);
    
    return (rcode == errSecSuccess);
}

bool get_codesign_status(bool insecure_result, bool secure_result) {
    printf("Safe Validation   : %s\n",   secure_result ? VERIFIED_STRN : FAILED_STRN);
    printf("Unsafe Validation : %s\n", insecure_result ? VERIFIED_STRN : FAILED_STRN);
    if (insecure_result && (!secure_result)) {
        printf("\nCodesign evasion detected\n");
        return ERR_SIGNING_EVASION;
    } else if ((!insecure_result) && (!secure_result)) {
        printf("\nBinary is not signed\n");
        return ERR_FILE_NOT_SIGNED;
    } else if ((!insecure_result) && secure_result) {
        printf("\nVery strange! secure codesign check reports valid but insecure codesigning does not\n");
        printf("Marking as evasion");
        return ERR_SIGNING_EVASION;
    }
    printf("\nSignature verified successfully\n");
    return ERR_VALID_SIGNATURE;
    
}

int main(int argc, const char * argv[]) {
    @autoreleasepool {
        bool anchor_trusted = false;
        const char* binary_path = argv[0];
        unsigned int path_arg = 1;
        if (argc >= 2) {
            if (0 == strncmp("-a", argv[1], 2)) {
                anchor_trusted = true;
                argc--;
                path_arg++;
            }
            if (0 == strncmp("-h",argv[1], 2)) {
                argc = 1;
            }
        }
        if (argc != 2) {
            printf("are-you-really-signed By SentinelOne\n");
            printf("Check if a binary attempts codesign evasion using Universal/FAT Binary technique discovered by Okta\n");
            printf("More details: https://www.okta.com/security-blog/2018/06/issues-around-third-party-apple-code-signing-checks/\n");
            
            printf("\nReturn Codes\n");
            printf("0 - Signed with valid chain validation (anchor apple)\n");
            printf("1 - Signing evasion detected\n");
            printf("2 - File not signed\n");
            printf("3 - File not found/Usage\n");
            
            printf("\nUsage: %s [-h] [-a] <Signed File>\n", binary_path);
            printf("Flags:\n");
            printf("\t-a: Check with 'anchor trusted' instead of 'anchor apple generic' - Pin to Cert Store instead of Apple Certs (for corporate purposes mostly).\n");
            printf("\t-h: Print help and usage\n");
            return ERR_FILE_NOT_FOUND_USAGE;
        }
        NSString* target_file = [NSString stringWithUTF8String :argv[path_arg]];
        NSFileManager *fileManager = [NSFileManager defaultManager];
        if (![fileManager fileExistsAtPath:target_file]) {
            printf("File %s not found!\n", argv[path_arg]);
            return ERR_FILE_NOT_SIGNED;
        }
        NSURL* target_file_url = [NSURL URLWithString:target_file];
        bool valid_insecure = insecure_validation(target_file_url);
        bool valid_secure   = secure_validation(target_file_url, anchor_trusted);
        return get_codesign_status(valid_insecure, valid_secure);
    }
    return 0;
}
