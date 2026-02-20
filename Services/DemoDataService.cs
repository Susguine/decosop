using DecoSOP.Data;
using DecoSOP.Models;
using Microsoft.EntityFrameworkCore;

namespace DecoSOP.Services;

public static class DemoDataService
{
    /// <summary>
    /// Returns true if data was seeded, false if skipped (data already exists).
    /// </summary>
    public static async Task<bool> SeedDemoDataAsync(AppDbContext db)
    {
        // Skip if any data already exists
        if (await db.Categories.AnyAsync() ||
            await db.SopCategories.AnyAsync() ||
            await db.DocumentCategories.AnyAsync() ||
            await db.WebDocCategories.AnyAsync())
            return false;

        // === SOP Categories & Documents ===

        var onboarding = new Category { Name = "Onboarding", SortOrder = 0, IsFavorited = true, Color = "blue" };
        var safety = new Category { Name = "Safety Procedures", SortOrder = 1, IsPinned = true, Color = "red" };
        var itPolicies = new Category { Name = "IT Policies", SortOrder = 2, Color = "purple" };
        var hr = new Category { Name = "HR", SortOrder = 3 };

        db.Categories.AddRange(onboarding, safety, itPolicies, hr);
        await db.SaveChangesAsync();

        // Onboarding SOPs
        db.Documents.AddRange(
            new SopDocument
            {
                CategoryId = onboarding.Id, Title = "New Employee Checklist", SortOrder = 0,
                HtmlContent = "<h1>New Employee Checklist</h1><h2>Before Day 1</h2><ul><li>Send welcome email with start date and parking info</li><li>Set up workstation and accounts</li><li>Prepare badge and access cards</li></ul><h2>Day 1</h2><ul><li>Facility tour</li><li>Meet the team introductions</li><li>Review company handbook</li><li>Complete HR paperwork</li></ul><h2>First Week</h2><ul><li>Department-specific training</li><li>Safety orientation</li><li>IT systems walkthrough</li></ul>"
            },
            new SopDocument
            {
                CategoryId = onboarding.Id, Title = "IT Account Setup Guide", SortOrder = 1, IsFavorited = true,
                HtmlContent = "<h1>IT Account Setup Guide</h1><h2>Required Accounts</h2><ol><li><strong>Email</strong> &mdash; Submit request to IT Help Desk with employee name and department</li><li><strong>Network Login</strong> &mdash; Created automatically from HR system</li><li><strong>VPN Access</strong> &mdash; Required for remote work; submit VPN request form</li></ol><h2>Software Installation</h2><p>Standard workstation image includes Office suite, PDF reader, and web browser. Additional software requires manager approval via the Software Request Form.</p>"
            },
            new SopDocument
            {
                CategoryId = onboarding.Id, Title = "Benefits Enrollment", SortOrder = 2,
                HtmlContent = "<h1>Benefits Enrollment</h1><p>New employees must complete benefits enrollment within <strong>30 days</strong> of their start date.</p><h2>Available Plans</h2><table><tr><th>Plan</th><th>Coverage</th><th>Employee Cost</th></tr><tr><td>Basic Health</td><td>Individual</td><td>$50/month</td></tr><tr><td>Family Health</td><td>Family</td><td>$150/month</td></tr><tr><td>Dental</td><td>Individual + Family</td><td>$25/month</td></tr></table><p>Contact HR for questions about plan details and dependent coverage.</p>"
            }
        );

        // Safety SOPs
        db.Documents.AddRange(
            new SopDocument
            {
                CategoryId = safety.Id, Title = "Fire Evacuation Procedure", SortOrder = 0,
                HtmlContent = "<h1>Fire Evacuation Procedure</h1><h2>When the Alarm Sounds</h2><ol><li>Stop all work immediately</li><li>Do NOT use elevators</li><li>Proceed to the nearest marked exit</li><li>Walk, do not run</li><li>Report to your designated assembly point</li></ol><h2>Assembly Points</h2><ul><li><strong>Building A</strong> &mdash; North parking lot</li><li><strong>Building B</strong> &mdash; South sidewalk near flagpole</li></ul><blockquote>Floor wardens are responsible for sweeping their assigned areas before evacuating.</blockquote>"
            },
            new SopDocument
            {
                CategoryId = safety.Id, Title = "First Aid Response", SortOrder = 1,
                HtmlContent = "<h1>First Aid Response</h1><h2>Minor Injuries</h2><p>First aid kits are located in every break room and near all exits. For minor cuts, burns, or scrapes, use the supplies in the nearest kit.</p><h2>Serious Injuries</h2><ol><li>Call 911 immediately</li><li>Notify a trained first responder (see posted list)</li><li>Do not move the injured person unless they are in immediate danger</li><li>Stay with the person until help arrives</li></ol><h2>Reporting</h2><p>All incidents must be reported using the <strong>Incident Report Form</strong> within 24 hours.</p>"
            }
        );

        // IT Policies
        db.Documents.AddRange(
            new SopDocument
            {
                CategoryId = itPolicies.Id, Title = "Password Policy", SortOrder = 0,
                HtmlContent = "<h1>Password Policy</h1><h2>Requirements</h2><ul><li>Minimum 12 characters</li><li>Must include uppercase, lowercase, number, and special character</li><li>Cannot reuse last 10 passwords</li><li>Must be changed every 90 days</li></ul><h2>Best Practices</h2><ul><li>Use a password manager</li><li>Never share passwords via email or chat</li><li>Enable multi-factor authentication where available</li></ul>"
            },
            new SopDocument
            {
                CategoryId = itPolicies.Id, Title = "Acceptable Use Policy", SortOrder = 1,
                HtmlContent = "<h1>Acceptable Use Policy</h1><p>Company technology resources are provided for business purposes. Limited personal use is permitted provided it does not interfere with work duties.</p><h2>Prohibited Activities</h2><ul><li>Installing unauthorized software</li><li>Accessing inappropriate or illegal content</li><li>Sharing confidential data outside approved channels</li><li>Using company resources for personal commercial gain</li></ul><p>Violations may result in disciplinary action up to and including termination.</p>"
            }
        );

        // HR
        db.Documents.Add(
            new SopDocument
            {
                CategoryId = hr.Id, Title = "Time Off Request Process", SortOrder = 0,
                HtmlContent = "<h1>Time Off Request Process</h1><h2>How to Request</h2><ol><li>Submit request through the HR portal at least 2 weeks in advance</li><li>Your manager will receive an email notification</li><li>Approval or denial will be communicated within 3 business days</li></ol><h2>Types of Leave</h2><ul><li><strong>Vacation</strong> &mdash; Accrued at 1.5 days/month</li><li><strong>Sick</strong> &mdash; 10 days per year</li><li><strong>Personal</strong> &mdash; 3 days per year</li></ul>"
            }
        );

        await db.SaveChangesAsync();

        // === SOP File Categories (empty, for uploaded SOP files) ===

        var procedures = new SopCategory { Name = "Procedures", SortOrder = 0, Color = "orange" };
        var formsTemplates = new SopCategory { Name = "Forms & Templates", SortOrder = 1, Color = "yellow" };
        var complianceDocs = new SopCategory { Name = "Compliance", SortOrder = 2, IsFavorited = true };

        db.SopCategories.AddRange(procedures, formsTemplates, complianceDocs);
        await db.SaveChangesAsync();

        // === Document Categories (empty, for uploaded files) ===

        var forms = new DocumentCategory { Name = "Forms", SortOrder = 0, Color = "green" };
        var templates = new DocumentCategory { Name = "Templates", SortOrder = 1, Color = "orange" };
        var references = new DocumentCategory { Name = "References", SortOrder = 2, IsFavorited = true };

        db.DocumentCategories.AddRange(forms, templates, references);
        await db.SaveChangesAsync();

        // === Web Doc Categories & Documents ===

        var trainingGuides = new WebDocCategory { Name = "Training Guides", SortOrder = 0, Color = "teal", IsFavorited = true };
        var quickRefs = new WebDocCategory { Name = "Quick References", SortOrder = 1, IsPinned = true };

        db.WebDocCategories.AddRange(trainingGuides, quickRefs);
        await db.SaveChangesAsync();

        db.WebDocuments.AddRange(
            new WebDocument
            {
                CategoryId = trainingGuides.Id, Title = "Email System Guide", SortOrder = 0,
                HtmlContent = "<h1>Email System Guide</h1><h2>Accessing Email</h2><p>Access your company email at <strong>mail.example.com</strong> or through the desktop Outlook client.</p><h2>Email Signature</h2><p>All employees must use the standard company email signature:</p><blockquote>Jane Smith<br/>Title | Department<br/>Company Name<br/>Phone: (555) 123-4567</blockquote><h2>Tips</h2><ul><li>Check email at least twice daily</li><li>Respond to internal messages within 24 hours</li><li>Use &ldquo;Reply All&rdquo; sparingly</li></ul>"
            },
            new WebDocument
            {
                CategoryId = trainingGuides.Id, Title = "Conference Room Booking", SortOrder = 1,
                HtmlContent = "<h1>Conference Room Booking</h1><h2>Available Rooms</h2><table><tr><th>Room</th><th>Capacity</th><th>Equipment</th></tr><tr><td>Room A (2nd Floor)</td><td>8</td><td>Projector, Whiteboard</td></tr><tr><td>Room B (2nd Floor)</td><td>4</td><td>TV Screen</td></tr><tr><td>Board Room (3rd Floor)</td><td>20</td><td>Projector, Video Conferencing</td></tr></table><h2>How to Book</h2><ol><li>Open the shared calendar in Outlook</li><li>Select the room and time slot</li><li>Add a descriptive title for your meeting</li><li>Send the invitation</li></ol>"
            },
            new WebDocument
            {
                CategoryId = quickRefs.Id, Title = "Wi-Fi Connection", SortOrder = 0, IsFavorited = true,
                HtmlContent = "<h1>Wi-Fi Connection</h1><h2>Corporate Network</h2><ul><li><strong>SSID:</strong> CorpNet-Secure</li><li><strong>Authentication:</strong> Use your network login credentials</li></ul><h2>Guest Network</h2><ul><li><strong>SSID:</strong> CorpNet-Guest</li><li><strong>Password:</strong> Posted at reception (rotates weekly)</li></ul><p>The guest network has limited bandwidth and no access to internal resources.</p>"
            },
            new WebDocument
            {
                CategoryId = quickRefs.Id, Title = "Printer Setup", SortOrder = 1,
                HtmlContent = "<h1>Printer Setup</h1><h2>Windows</h2><ol><li>Open <strong>Settings &gt; Devices &gt; Printers</strong></li><li>Click <strong>Add a printer</strong></li><li>Select the printer for your floor from the list</li></ol><h2>Printer Locations</h2><ul><li><strong>1st Floor:</strong> HP-1F-Color (color) &amp; HP-1F-BW (black and white)</li><li><strong>2nd Floor:</strong> HP-2F-Color</li><li><strong>3rd Floor:</strong> HP-3F-BW</li></ul><p>For toner replacements or paper jams, contact Facilities at ext. 4200.</p>"
            }
        );

        await db.SaveChangesAsync();

        return true;
    }
}
