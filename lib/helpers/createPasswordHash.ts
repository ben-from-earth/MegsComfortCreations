import bcrypt from 'bcrypt';

async function main() {
  const rawPassword = 'MCC1097!'; // <-- change this
  const salt = await bcrypt.genSalt(); // uses default rounds (10)
  const hash = await bcrypt.hash(rawPassword, salt);

  console.log('Hash:', hash);
}

main().catch(console.error);
