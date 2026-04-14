import { getUncachableStripeClient } from './stripeClient';

async function createOrUpdateProducts() {
  const stripe = await getUncachableStripeClient();

  const NEW_NAME = 'Microsoft 365 Connector Pro';
  const NEW_DESC = 'Connect your Microsoft 365 email, calendar, OneDrive, and Teams to Claude.ai via MCP. Read, search, and summarize your work data directly in Claude conversations.';

  // Check for existing product (old or new name)
  const [existingNew, existingOld] = await Promise.all([
    stripe.products.search({ query: `name:'${NEW_NAME}'` }),
    stripe.products.search({ query: "name:'Outlook Connector Pro'" }),
  ]);

  if (existingNew.data.length > 0) {
    console.log('Product already up-to-date:', existingNew.data[0].id);
    const prices = await stripe.prices.list({ product: existingNew.data[0].id, active: true });
    console.log('Prices:', prices.data.map(p => `${p.id}: $${(p.unit_amount || 0) / 100}/${p.recurring?.interval}`));
    return;
  }

  if (existingOld.data.length > 0) {
    const product = existingOld.data[0];
    console.log('Updating existing product:', product.id);
    await stripe.products.update(product.id, {
      name: NEW_NAME,
      description: NEW_DESC,
    });
    console.log('Product updated to:', NEW_NAME);
    const prices = await stripe.prices.list({ product: product.id, active: true });
    console.log('Prices:', prices.data.map(p => `${p.id}: $${(p.unit_amount || 0) / 100}/${p.recurring?.interval}`));
    return;
  }

  // Create new product
  const product = await stripe.products.create({
    name: NEW_NAME,
    description: NEW_DESC,
    metadata: { type: 'subscription' },
  });

  const monthlyPrice = await stripe.prices.create({
    product: product.id,
    unit_amount: 999,
    currency: 'usd',
    recurring: { interval: 'month' },
  });

  const yearlyPrice = await stripe.prices.create({
    product: product.id,
    unit_amount: 7999,
    currency: 'usd',
    recurring: { interval: 'year' },
  });

  console.log('Created product:', product.id);
  console.log('Monthly price:', monthlyPrice.id, '- $9.99/month');
  console.log('Yearly price:', yearlyPrice.id, '- $79.99/year');
}

createOrUpdateProducts().catch(console.error);
