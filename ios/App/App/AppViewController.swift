import UIKit
import Capacitor

// In-app Capacitor plugin exposing whether the app currently runs as a
// resizable window (iPadOS 26 windowing / Stage Manager) rather than
// fullscreen. The web layer uses it to keep the title bar clear of the
// window's traffic-light controls and the footer clear of the rounded corners
// and resize grip — regions iPadOS does not (reliably) report through the
// webview's safe-area insets.
@objc(WindowStatePlugin)
public class WindowStatePlugin: CAPPlugin, CAPBridgedPlugin {
    public let identifier = "WindowStatePlugin"
    public let jsName = "WindowState"
    public let pluginMethods: [CAPPluginMethod] = [
        CAPPluginMethod(name: "get", returnType: CAPPluginReturnPromise)
    ]

    fileprivate var windowed = false

    // Called by the web layer on every page load (the settings/help pages are
    // separate documents, so each one pulls the current state on startup).
    @objc func get(_ call: CAPPluginCall) {
        call.resolve(["windowed": windowed])
    }

    fileprivate func push(_ value: Bool) {
        windowed = value
        notifyListeners("windowedchange", data: ["windowed": value])
    }
}

class AppViewController: CAPBridgeViewController {
    private let windowState = WindowStatePlugin()
    private var lastWindowed: Bool?

    override open func capacitorDidLoad() {
        bridge?.registerPluginInstance(windowState)
    }

    override func viewDidLayoutSubviews() {
        super.viewDidLayoutSubviews()
        syncWindowedState()
    }

    override func viewSafeAreaInsetsDidChange() {
        super.viewSafeAreaInsetsDidChange()
        syncWindowedState()
    }

    // Windowed = the app's window is meaningfully smaller than its screen.
    // Fullscreen iPad and all iPhones compare equal (the keyboard resizes the
    // webview, not the window, so it never affects this).
    private func syncWindowedState() {
        guard let window = view.window, let screen = window.windowScene?.screen else { return }
        let wb = window.bounds.size
        let sb = screen.bounds.size
        let windowed = (sb.width - wb.width) > 1 || (sb.height - wb.height) > 1
        if windowed != lastWindowed {
            lastWindowed = windowed
            windowState.push(windowed)
        }
    }
}
