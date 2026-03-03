const { SPFetchClient } = require('@pnp/nodejs');
const { sp } = require('@pnp/sp/presets/all');

async function main() {
    sp.setup({
        sp: {
            fetchClientFactory: () => new SPFetchClient(
                'https://barlogpl.sharepoint.com/sites/ta-management',
                process.env.SP_ID,
                process.env.SP_SECRET
            ),
        },
    });

    try {
        const list = sp.web.lists.getByTitle('Projekte');
        const fields = await list.fields.filter("Hidden eq false and ReadOnlyField eq false").get();

        console.log("=== Felder in der Projekte Liste ===");
        fields.forEach(f => {
            console.log(`- ${f.Title} (${f.InternalName})`);
        });
    } catch (error) {
        console.error("Error:", error);
    }
}

main();
