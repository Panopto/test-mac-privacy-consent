//
//  PrivacyConsentController.h
//  Panopto Recorder
//
//  Created by Hal Mueller on 8/20/18.
//

#import <Foundation/Foundation.h>

NS_ASSUME_NONNULL_BEGIN

// we need our own Enum because the system's AVAuthorizationStatus is not available prior to 10.14
typedef NS_ENUM(NSInteger, PrivacyConsentState) {
    PrivacyConsentStateUnknown,
    PrivacyConsentStateGranted,
    PrivacyConsentStateDenied
};

/// Object to broker privacy consent requests and report status of access for camera, microphone, and AppleScript.
@interface PrivacyConsentController : NSObject
+ (instancetype)sharedController;

/// This is the best single entry point to PrivacyConsentController. Fires permissions request for camera and microphone,
/// optionally throws up nag dialogs if appropriate. Completion handler runs on main thread.
- (void)requestMediaConsentNagIfDenied:(BOOL)nagIfDenied completion:(void (^)(BOOL))allMediaAccessGranted;

- (PrivacyConsentState)automationConsentForKeynotePromptIfNeeded:(BOOL)promptIfNeeded;
- (PrivacyConsentState)automationConsentForPowerPointPromptIfNeeded:(BOOL)promptIfNeeded;
- (PrivacyConsentState)automationConsentForBundleIdentifier:(NSString *)bundleIdentifier promptIfNeeded:(BOOL)promptIfNeeded;

- (BOOL)hasMicrophoneConsent;
- (BOOL)hasCameraConsent;
- (BOOL)hasFullMediaConsent;

- (void)nagForMicrophoneConsentIfNeeded;
- (void)nagForCameraConsentIfNeeded;
- (void)nagForKeynoteConsent;
- (void)nagForPowerPointConsent;

@end

NS_ASSUME_NONNULL_END
