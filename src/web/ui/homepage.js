export function setupHomepage(mainContent) {
  mainContent.innerHTML = `
    <section class="modern-container-text" style="--container-max-width: 900px;">
      <h1 class="homepage-title">Welcome to Free Excel Generators</h1>
      <p class="homepage-intro-text">
        I'm just a real human—no AI startup, no SaaS company, no hidden subscriptions. I built this site as a personal project to sharpen my skills in JavaScript, HTML, and Excel, while also creating tools that actually help people. I've always been frustrated by websites that require logins just to download a spreadsheet, or by AI tools that either cost too much, have a steep learning curve, or don't quite do what you need. This project is my way of pushing back on all that. It's made for the folks stuck in admin roles or team support positions who never had time to conquer Excel but still want clean, functional spreadsheets without the headache. Maybe you just hate tools altogether. Maybe you want to impress your boss with a polished calendar or tracker and wink wink pretend you built it—go for it, I won't stop you. Everything here is free, built to work entirely in your browser, and focused on being genuinely useful. If it helps you out, maybe consider buying me a coffee—or better yet, shoot me an email telling me how much you love the site. Who knows, you might just show up on the future testimonial wall. Thanks for being here
      </p>
      <p class="homepage-intro-contact">
        If you want something custom or hand-tailored to your workflow, I do take custom orders—just send me an email. We can set up a time to talk it through, or hash it out the old-fashioned way: through the world's fastest postal service—email. Got a suggestion? Want to send me praise, hate mail, or just test if I actually read my inbox? Maybe you're a scam bot who wandered in by accident. Either way, I'm all ears. Shoot me a message—I read everything.
      </p>
    </section>
    <section class="modern-container-tool" style="--container-max-width: 750px; text-align: center;">
      <h3 class="homepage-updates-title">🔧 What's Coming Next</h3>
      <div style="display: flex; flex-direction: column; align-items: center; gap: 8px; margin: 18px 0 18px 0; font-size: 1.15rem;">
        <div style="display: flex; align-items: center; gap: 8px;"><span>🆕</span><span>Attendance Tracker (coming soon!)</span></div>
        <div style="display: flex; align-items: center; gap: 8px;"><span>✨</span><span>Improved styling and theme options</span></div>
        <div style="display: flex; align-items: center; gap: 8px;"><span>📩</span><span>User-suggested features for all tools</span></div>
      </div>
      <p class="homepage-updates-contact">
        Got a feature idea or tool request? <a href="#" id="suggestLink" class="homepage-suggest-link">Submit it here</a>.
      </p>
    </section>
  `;
  setTimeout(() => {
    const suggestLink = document.getElementById('suggestLink');
    if (suggestLink) {
      suggestLink.addEventListener('click', function(e) {
        e.preventDefault();
        const navContact = document.getElementById('nav-contact');
        if (navContact) navContact.click();
      });
    }
  }, 100);
} 