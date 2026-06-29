export default function Home() {
  return (
    <main className="min-h-screen bg-zinc-950 text-white flex flex-col">
      <header className="border-b border-zinc-800 p-5">
        <div className="max-w-5xl mx-auto flex items-center justify-between">
          <div>
            <h1 className="text-3xl font-bold">JARVIS</h1>
            <p className="text-sm text-zinc-400">
              Personal AI Operating System
            </p>
          </div>

          <div className="flex items-center gap-2">
            <div className="h-3 w-3 rounded-full bg-green-500 animate-pulse"></div>
            <span className="text-sm text-zinc-400">ONLINE</span>
          </div>
        </div>
      </header>

      <section className="flex-1 flex items-center justify-center">
        <div className="max-w-3xl w-full px-6">
          <h2 className="text-4xl font-bold mb-4">
            Good Afternoon, Curt.
          </h2>

          <p className="text-zinc-400 text-lg mb-10">
            Awaiting your instructions...
          </p>

          <div className="rounded-xl border border-zinc-800 bg-zinc-900 p-4">
            <input
              type="text"
              placeholder="Type a command..."
              className="w-full bg-transparent outline-none text-lg placeholder:text-zinc-500"
            />
          </div>

          <button className="mt-6 rounded-lg bg-blue-600 px-6 py-3 font-semibold hover:bg-blue-500 transition">
            Send
          </button>
        </div>
      </section>
    </main>
  );
}