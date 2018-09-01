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
        NSLog(@"granted: %d", granted);
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
