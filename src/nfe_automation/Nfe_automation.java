package nfe_automation;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxProfile;
import java.util.concurrent.TimeUnit;
import org.openqa.selenium.firefox.internal.ProfilesIni;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;

public class Nfe_automation {

    public static void main(String[] args) throws Exception {
        ProfilesIni profilesIni = new ProfilesIni();
        FirefoxProfile profile = profilesIni.getProfile("default");
        WebDriver driver = new FirefoxDriver(profile);

        DesiredCapabilities capabilities = new DesiredCapabilities();
        capabilities.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
        
        profile.setPreference("browser.ssl_override_behavior", 1);
        profile.setPreference("network.proxy.type", 0);
        profile.setAcceptUntrustedCertificates(true);
        profile.setAssumeUntrustedCertificateIssuer(false);

        driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);

        OpenUrl Openurl = new OpenUrl();
        Openurl.OpenURL(driver);
        
        driver.close();
        driver.quit();
    }
}
