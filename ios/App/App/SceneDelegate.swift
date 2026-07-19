import UIKit

// UIScene lifecycle adoption. Without a scene manifest, iPadOS 26 treats the
// app as legacy and reserves a system "title bar" strip at the top of every
// window (rendered as a black band above the web content). With scenes, the
// app's content extends to the top of the window and the window controls
// contribute to the safe area instead. UIKit also requires scene adoption for
// apps built with the next SDK, so this is the forward path regardless.
//
// The window and root view controller come from Main.storyboard
// (UISceneStoryboardFile in Info.plist); UIKit assigns `window` before
// calling scene(_:willConnectTo:options:).
class SceneDelegate: UIResponder, UIWindowSceneDelegate {
    var window: UIWindow?

    func scene(_ scene: UIScene, willConnectTo session: UISceneSession, options connectionOptions: UIScene.ConnectionOptions) {
        // Theme the native window surface so no system-black background can
        // ever peek through around the webview (launch, rotation, resize).
        window?.backgroundColor = UIColor.systemBackground
    }
}
