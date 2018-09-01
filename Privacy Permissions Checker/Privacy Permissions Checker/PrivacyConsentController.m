//
//  PrivacyConsentController.m
//  Panopto Recorder
//
//  Created by Hal Mueller on 8/20/18.
//

// Privacy consent rules changed as of macOS 10.14. Prior to 10.14, applications always had access to camera and microphone.
// The "if (@available(macOS 10.14, *))" conditional executions reflect that.
//
// For testing, you can use "tccutil" to reset permissions state:
//     tccutil reset All
//     tccutil reset Camera
//     tccutil reset Microphone
//     tccutil reset AppleEvents

#import "PrivacyConsentController.h"
#import <AVFoundation/AVFoundation.h>
#import <Cocoa/Cocoa.h>

@interface PrivacyConsentController ()

@property(atomic, assign) PrivacyConsentState microphoneConsentState;
@property(atomic, assign) PrivacyConsentState cameraConsentState;
@end

@implementation PrivacyConsentController

NSString *keynoteBundleIdentifier = @"com.apple.iWork.Keynote";
NSString *powerPointBundleIdentifier = @"com.microsoft.PowerPoint";

+ (instancetype)sharedController
{
    static dispatch_once_t once;
    static PrivacyConsentController *sharedController;
    dispatch_once(&once, ^{
        sharedController = [[self alloc] init];
    });
    return sharedController;
}

- (instancetype)init
{
    if (self = [super init]) {
        if (@available(macOS 10.14, *)) {
            // case AVAuthorizationStatusRestricted means the device doesn't exist or isn't available. We treat this as "granted", no reason to block recording or pester for permission. Media layer can block later if appropriate.
            switch ([AVCaptureDevice authorizationStatusForMediaType:AVMediaTypeVideo]) {
                case AVAuthorizationStatusAuthorized:
                case AVAuthorizationStatusRestricted:
                    _cameraConsentState = PrivacyConsentStateGranted;
                case AVAuthorizationStatusDenied:
                    _cameraConsentState = PrivacyConsentStateDenied;
                case AVAuthorizationStatusNotDetermined:
                    _cameraConsentState = PrivacyConsentStateUnknown;
            }
            switch ([AVCaptureDevice authorizationStatusForMediaType:AVMediaTypeAudio]) {
                case AVAuthorizationStatusAuthorized:
                case AVAuthorizationStatusRestricted:
                    _microphoneConsentState = PrivacyConsentStateGranted;
                case AVAuthorizationStatusDenied:
                    _microphoneConsentState = PrivacyConsentStateDenied;
                case AVAuthorizationStatusNotDetermined:
                    _microphoneConsentState = PrivacyConsentStateUnknown;
            }
            [[[NSWorkspace sharedWorkspace] notificationCenter] addObserver:self
                                                                   selector:@selector(applicationLaunched:)
                                                                       name:NSWorkspaceDidLaunchApplicationNotification
                                                                     object:nil];
        }
        else {
            _cameraConsentState = PrivacyConsentStateGranted;
            _microphoneConsentState = PrivacyConsentStateGranted;
        }
    }
    return self;
}

// If user hasn't clicked either option in a privacy consent dialog box, or if we have never asked, consent status is unknown.
// The requestAccessForMediaType:completionHandler: completion block is executed pretty much instantly if user has EVER seen this consent challenge before. If the dialog remains on screen, though, the completion handler won't fire until user says yay or nay.
- (BOOL)mediaConsentStatusKnown
{
    return ((self.microphoneConsentState != PrivacyConsentStateUnknown) &&
            (self.cameraConsentState != PrivacyConsentStateUnknown));
}

- (PrivacyConsentState)automationConsentForKeynotePromptIfNeeded:(BOOL)promptIfNeeded
{
    return [self automationConsentForBundleIdentifier:keynoteBundleIdentifier promptIfNeeded:promptIfNeeded];
}

- (PrivacyConsentState)automationConsentForPowerPointPromptIfNeeded:(BOOL)promptIfNeeded
{
    return [self automationConsentForBundleIdentifier:powerPointBundleIdentifier promptIfNeeded:promptIfNeeded];
}

// !!!: Workaround for Apple bug. Their AppleEvents.h header conditionally defines errAEEventWouldRequireUserConsent and one other constant, valid only for 10.14 and higher, which means our code inside the @available() check would fail to compile. Remove this definition when they fix it.
#if __MAC_OS_X_VERSION_MIN_REQUIRED <= __MAC_10_14
enum {
    errAEEventWouldRequireUserConsent = -1744, /* Determining whether this can be sent would require prompting the user, and the AppleEvent was sent with kAEDoNotPromptForPermission */
};
#endif

// Be careful about calling this method from the main thread, because AEDeterminePermissionToAutomateTarget() is a blocking call.
- (PrivacyConsentState)automationConsentForBundleIdentifier:(NSString *)bundleIdentifier promptIfNeeded:(BOOL)promptIfNeeded
{
    PrivacyConsentState result;
    if (@available(macOS 10.14, *)) {
        AEAddressDesc addressDesc;
        // We need a C string here, not an NSString
        const char *bundleIdentifierCString = [bundleIdentifier cStringUsingEncoding:NSUTF8StringEncoding];
        OSErr createDescResult = AECreateDesc(typeApplicationBundleID, bundleIdentifierCString, strlen(bundleIdentifierCString), &addressDesc);
        OSStatus appleScriptPermission = AEDeterminePermissionToAutomateTarget(&addressDesc, typeWildCard, typeWildCard, promptIfNeeded);
        AEDisposeDesc(&addressDesc);
        switch (appleScriptPermission) {
            case errAEEventWouldRequireUserConsent:
                NSLog(@"Automation consent not yet granted for %@, would require user consent.", bundleIdentifier);
                result = PrivacyConsentStateUnknown;
                break;
            case noErr:
                NSLog(@"Automation permitted for %@.", bundleIdentifier);
                result = PrivacyConsentStateGranted;
                break;
            case errAEEventNotPermitted:
                NSLog(@"Automation of %@ not permitted.", bundleIdentifier);
                result = PrivacyConsentStateDenied;
                break;
            case procNotFound:
                NSLog(@"%@ not running, automation consent unknown.", bundleIdentifier);
                result = PrivacyConsentStateUnknown;
                break;
            default:
                NSLog(@"%s switch statement fell through: %@ %d", __PRETTY_FUNCTION__, bundleIdentifier, appleScriptPermission);
                result = PrivacyConsentStateUnknown;
        }
        return result;
    }
    else {
        return PrivacyConsentStateGranted;
    }
}

- (void)requestMediaConsentNagIfDenied:(BOOL)nagIfDenied completion:(void (^)(BOOL))allMediaAccessGranted
{
    if (@available(macOS 10.14, *)) {
        [AVCaptureDevice requestAccessForMediaType:AVMediaTypeAudio completionHandler:^(BOOL granted) {
            if (granted) {
                self.microphoneConsentState = PrivacyConsentStateGranted;
            }
            else {
                self.microphoneConsentState = PrivacyConsentStateDenied;
            }
            [AVCaptureDevice requestAccessForMediaType:AVMediaTypeVideo completionHandler:^(BOOL granted) {
                if (granted) {
                    self.cameraConsentState = PrivacyConsentStateGranted;
                }
                else {
                    self.cameraConsentState = PrivacyConsentStateDenied;
                }
                if (nagIfDenied) {
                    dispatch_async(dispatch_get_main_queue(), ^{
                        [self nagForMicrophoneConsentIfNeeded];
                        [self nagForCameraConsentIfNeeded];
                    });
                }
                dispatch_async(dispatch_get_main_queue(), ^{
                    allMediaAccessGranted(self.hasFullMediaConsent);
                });
            }];
        }];
    }
    else {
        allMediaAccessGranted(self.hasFullMediaConsent);
    }
}

- (BOOL)hasMicrophoneConsent
{
    if (@available(macOS 10.14, *)) {
        return (self.microphoneConsentState == PrivacyConsentStateGranted);
    }
    // prior to macOS 10.14, we always have media consent
    return YES;
}

- (BOOL)hasCameraConsent
{
    if (@available(macOS 10.14, *)) {
        return (self.cameraConsentState == PrivacyConsentStateGranted);
    }
    // prior to macOS 10.14, we always have media consent
    return YES;
}

/// Return YES if user has been asked for consent to access to both Camera and Microphone (via system dialog that we triggered), has allowed that access, and has not subsequently withdrawn that consent (via tccutil or the Security & Privacy preference pane).
/// !!!: August 2018, I tried to allow session recording with audio + screen capture, no camera usage, no camera consent. Result was very unstable behavior. Therefore, DO NOT ALLOW RECORDING unless consent granted for both audio and video.
- (BOOL)hasFullMediaConsent
{
    return (self.hasCameraConsent && self.hasMicrophoneConsent);
}

- (void)nagForMicrophoneConsentIfNeeded
{
    if (self.microphoneConsentState == PrivacyConsentStateDenied) {
        NSAlert *alert = [[NSAlert alloc] init];
        alert.alertStyle = NSAlertStyleWarning;
        alert.messageText = @"Panopto needs access to the microphone";
        alert.informativeText = @"Panopto can't make recordings unless you grant permission for access to your microphone.";
        [alert addButtonWithTitle:@"Change Security & Privacy Preferences"];
        [alert addButtonWithTitle:@"Cancel"];
        
        NSInteger modalResponse = [alert runModal];
        if (modalResponse == NSAlertFirstButtonReturn) {
            [self launchPrivacyAndSecurityPreferencesMicrophoneSubPane];
        }
    }
}

- (void)nagForCameraConsentIfNeeded
{
    if (self.cameraConsentState == PrivacyConsentStateDenied) {
        NSAlert *alert = [[NSAlert alloc] init];
        alert.alertStyle = NSAlertStyleWarning;
        alert.messageText = @"Panopto needs access to the camera";
        alert.informativeText = @"Panopto can't make recordings unless you grant permission for access to your camera. We need this for either or both of screen capture and webcam capture.";
        [alert addButtonWithTitle:@"Change Security & Privacy Preferences"];
        [alert addButtonWithTitle:@"Cancel"];
        
        NSInteger modalResponse = [alert runModal];
        if (modalResponse == NSAlertFirstButtonReturn) {
            [self launchPrivacyAndSecurityPreferencesCameraSubPane];
        }
    }
}

- (void)nagForKeynoteConsent
{
    NSAlert *alert = [[NSAlert alloc] init];
    alert.alertStyle = NSAlertStyleWarning;
    alert.messageText = @"Panopto needs access to Keynote";
    alert.informativeText = @"Panopto can't record with these settings unless you grant permission for Automation control of Keynote.\n\nPlease change the setting and then restart Panopto.";
    [alert addButtonWithTitle:@"Change Security & Privacy Preferences"];
    [alert addButtonWithTitle:@"Cancel"];
    
    NSInteger modalResponse = [alert runModal];
    if (modalResponse == NSAlertFirstButtonReturn) {
        [self launchPrivacyAndSecurityPreferencesAutomationSubPane];
    }
}

- (void)nagForPowerPointConsent
{
    NSAlert *alert = [[NSAlert alloc] init];
    alert.alertStyle = NSAlertStyleWarning;
    alert.messageText = @"Panopto needs access to PowerPoint";
    alert.informativeText = @"Panopto can't record with these settings unless you grant permission for Automation control of PowerPoint.\n\nPlease change the setting and then restart Panopto.";
    [alert addButtonWithTitle:@"Change Security & Privacy Preferences"];
    [alert addButtonWithTitle:@"Cancel"];
    
    NSInteger modalResponse = [alert runModal];
    if (modalResponse == NSAlertFirstButtonReturn) {
        [self launchPrivacyAndSecurityPreferencesAutomationSubPane];
    }
}

// starting with https://macosxautomation.com/system-prefs-links.html, worked these links out experimentally 16 Aug 2018 with macOS 10.14.b6.
// Generic: [[NSWorkspace sharedWorkspace] openURL:[NSURL URLWithString:@"x-apple.systempreferences:com.apple.preference.security"]];
//    will open the Privacy & Security pane to whichever tab/subpane was last viewed.
- (void)launchPrivacyAndSecurityPreferencesMicrophoneSubPane
{
    [[NSWorkspace sharedWorkspace] openURL:[NSURL URLWithString:@"x-apple.systempreferences:com.apple.preference.security?Privacy_Microphone"]];
}

- (void)launchPrivacyAndSecurityPreferencesCameraSubPane
{
    [[NSWorkspace sharedWorkspace] openURL:[NSURL URLWithString:@"x-apple.systempreferences:com.apple.preference.security?Privacy_Camera"]];
}

- (void)launchPrivacyAndSecurityPreferencesAutomationSubPane
{
    // !!!: as of August 2018, toggling these settings in the pref pane DOES NOT warn you to restart the app to have changes take effect.
    [[NSWorkspace sharedWorkspace] openURL:[NSURL URLWithString:@"x-apple.systempreferences:com.apple.preference.security?Privacy_Automation"]];
}

- (void)applicationLaunched:(NSNotification *)notification
{
    NSRunningApplication *app = notification.userInfo[NSWorkspaceApplicationKey];
    if ([app.bundleIdentifier isEqualToString:keynoteBundleIdentifier]) {
        PrivacyConsentState keynoteConsent = [self automationConsentForKeynotePromptIfNeeded:YES];
        if (keynoteConsent != PrivacyConsentStateGranted) {
            [self nagForKeynoteConsent];
        }
    }
    else if ([app.bundleIdentifier isEqualToString:powerPointBundleIdentifier]) {
        PrivacyConsentState powerPointConsent = [self automationConsentForPowerPointPromptIfNeeded:YES];
        if (powerPointConsent != PrivacyConsentStateGranted) {
            [self nagForPowerPointConsent];
        }
    }
}
@end
