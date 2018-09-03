//
//  ViewController.m
//  Privacy Permissions Checker
//
//  Created by Hal Mueller on 8/27/18.
//  Copyright © 2018 Panopto. All rights reserved.
//

#import <Foundation/Foundation.h>
#import <CoreServices/CoreServices.h>

#import "ViewController.h"
#import "AppleScriptExerciser.h"
#import "PrivacyConsentController.h"

@interface ViewController ()
@property (strong) IBOutlet AppleScriptExerciser *appleScriptExerciser;
@end

@implementation ViewController

- (void)viewDidLoad {
    [super viewDidLoad];

    // Do any additional setup after loading the view.
}

- (IBAction)tellPowerPoint:(id)sender
{
    [self.appleScriptExerciser tellPowerPoint];
}

- (IBAction)showPowerPointVersion:(id)sender
{
    [self.appleScriptExerciser showPowerPointVersion];
}

- (IBAction)tellKeynote:(id)sender
{
    [self.appleScriptExerciser tellKeynote];
}

- (IBAction)showKeynoteVersion:(id)sender
{
    [self.appleScriptExerciser showKeynoteVersion];
}

- (IBAction)keynoteApplication:(id)sender
{
    NSLog(@"%@", self.appleScriptExerciser.keynoteApplication);
}

- (IBAction)powerPointApplication:(id)sender
{
    NSLog(@"%@", self.appleScriptExerciser.powerPointApplication);
}

- (IBAction)requestMediaConsent:(id)sender
{
    [[PrivacyConsentController sharedController] requestMediaConsentNagIfDenied:YES completion:^(BOOL granted) {
        // For demonstration purposes, we only log the result. But you can use this block to load another view controller/perform a segue, with different behavior depending on whether you have consent or not
        NSLog(@"granted: %d", granted);
        if (granted) {
            NSLog(@"proceed to the protected content that requires camera/microphone access");
        }
        else {
            NSLog(@"present alternate UI/content because the user denied consent for camera/microphone access");
        }
    }];
}

- (IBAction)nagForMediaConsent:(id)sender
{
    [[PrivacyConsentController sharedController] nagForCameraConsentIfNeeded];
    [[PrivacyConsentController sharedController] nagForMicrophoneConsentIfNeeded];
}

- (IBAction)determinePowerPointPermission:(id)sender
{
    PrivacyConsentState powerPointConsent = [[PrivacyConsentController sharedController] automationConsentForPowerPointPromptIfNeeded:YES];
    NSLog(@"PowerPoint consent %ld", powerPointConsent);
}

- (IBAction)determineKeynotePermission:(id)sender
{
    PrivacyConsentState keynoteConsent = [[PrivacyConsentController sharedController] automationConsentForKeynotePromptIfNeeded:YES];
    NSLog(@"Keynote consent %ld", keynoteConsent);
}

- (IBAction)requestPowerPointPermissionAndNag:(id)sender
{
    PrivacyConsentState powerPointConsent = [[PrivacyConsentController sharedController] automationConsentForPowerPointPromptIfNeeded:YES];
    if (powerPointConsent != PrivacyConsentStateGranted) {
        [[PrivacyConsentController sharedController] nagForPowerPointConsent];
    }
}

- (IBAction)requestKeynotePermissionAndNag:(id)sender
{
    PrivacyConsentState keynoteConsent = [[PrivacyConsentController sharedController] automationConsentForKeynotePromptIfNeeded:YES];
    if (keynoteConsent != PrivacyConsentStateGranted) {
        [[PrivacyConsentController sharedController] nagForKeynoteConsent];
    }
}

@end
