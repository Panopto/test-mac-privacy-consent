//
//  AppleScriptExerciser.m
//  Privacy Permissions Checker
//
//  Created by Hal Mueller on 8/27/18.
//  Copyright © 2018 Panopto. All rights reserved.
//

#import "AppleScriptExerciser.h"
#import <Cocoa/Cocoa.h>
#import "Keynote.h"
#import "PowerPoint.h"

@interface AppleScriptExerciser ()
@property (nonatomic, strong) KeynoteApplication *keynoteApplication;
@property (nonatomic, strong) PowerPointApplication *powerPointApplication;

/// Undocumented but apparently functional scripting property, but this works as of iWork 6.5
@property (readonly) BOOL playing;

@end

@implementation AppleScriptExerciser

- (instancetype)init
{
    if (self = [super init]) {
        _keynoteApplication = (KeynoteApplication *) [SBApplication applicationWithBundleIdentifier:@"com.apple.iWork.Keynote"];
        _keynoteApplication.delegate = self;
        _powerPointApplication = (PowerPointApplication *)[SBApplication applicationWithBundleIdentifier:@"com.microsoft.PowerPoint"];
        _powerPointApplication.delegate = self;
    }
    return self;
}

- (void)tellPowerPoint
{
    // excerpted from PowerPointCapture.m savePresentationAs:
    
    // If sandboxing is turned ON, this code fails with
    //NSAppleScriptErrorMessage = "Microsoft PowerPoint got an error: Application isn\U2019t running.";
    //NSAppleScriptErrorNumber = "-600";
    // Need to have permission? This script works in AppleScript.
    NSAppleScript* scriptObject = [[NSAppleScript alloc] initWithSource:
                                   [NSString stringWithFormat:
                                    @"tell application \"Microsoft PowerPoint\"\n"
                                    @"    activate\n"
                                    @"if active presentation exists then\n"
                                    @"set aPres to active presentation\n"
                                    @"else\n"
                                    @"set aPres to make new presentation\n"
                                    @"end if\n"
                                    @"set newSlide to make new slide at the end of aPres\n"
                                    @"end tell"]];
    NSDictionary *errorDict = nil;
    NSAppleEventDescriptor *executeResult = [scriptObject executeAndReturnError:&errorDict];
    if (executeResult) {
        NSLog(@"%@", executeResult);
    }
    else {
        NSBeep();
        NSLog(@"error %@", errorDict);
    }
}

- (void)showPowerPointVersion
{
    // extracted from PowerPointCapture.m
    PowerPointApplication *powerPoint = (PowerPointApplication *)[SBApplication applicationWithBundleIdentifier:@"com.microsoft.PowerPoint"];
    NSLog(@"PowerPoint %@", powerPoint);
    
    static NSString *nullString = @"(null)";
    if (powerPoint.Version == nullString)
    {
        // Scripting is not ready to get the version. Skip it.
        return;
    }
    
    NSString *powerPointVersion = powerPoint.Version;
    NSLog(@"PowerPointCapture target version = %@", powerPointVersion);
}

- (void)tellKeynote
{
    NSAppleScript* scriptObject = [[NSAppleScript alloc] initWithSource:
                                   [NSString stringWithFormat:
                                    @"tell application \"Keynote\"\n"
                                    @"    activate\n"
                                    @"end tell"]];
    NSDictionary *errorDict = nil;
    NSAppleEventDescriptor *executeResult = [scriptObject executeAndReturnError:&errorDict];
    if (executeResult) {
        NSLog(@"%@", executeResult);
    }
    else {
        NSBeep();
        NSLog(@"error %@", errorDict);
    }
}

- (void)showKeynoteVersion
{
    NSLog(@"Keynote %@", self.keynoteApplication);
    NSString *keynoteVersion = self.keynoteApplication.version;
    NSLog(@"Keynote version strng %@", keynoteVersion);
}

#pragma mark - SBApplicationDelegate
- (nullable id) eventDidFail:(const AppleEvent *)event withError:(NSError *)error
{
    NSLog(@"%s %@", __PRETTY_FUNCTION__, error);
    return nil;
}
@end
