//
//  AppleScriptExerciser.h
//  Privacy Permissions Checker
//
//  Created by Hal Mueller on 8/27/18.
//  Copyright © 2018 Panopto. All rights reserved.
//

#import <Foundation/Foundation.h>
#import <ScriptingBridge/ScriptingBridge.h>

NS_ASSUME_NONNULL_BEGIN
@class KeynoteApplication;
@class PowerPointApplication;

@interface AppleScriptExerciser : NSObject <SBApplicationDelegate>
@property (nonatomic, readonly) KeynoteApplication *keynoteApplication;
@property (nonatomic, readonly) PowerPointApplication *powerPointApplication;

- (void)tellPowerPoint;
- (void)showPowerPointVersion;
- (void)tellKeynote;
- (void)showKeynoteVersion;

@end

NS_ASSUME_NONNULL_END
