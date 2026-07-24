import React, { useEffect, useRef } from "react";

// Latar animasi untuk halaman login: partikel yang bergerak pelan dengan garis
// penghubung (efek "jaringan"), formula & warna disamakan dengan kode referensi,
// ditambah beberapa kubus 3D tembus pandang yang berputar dan melayang.
export default function LoginAnimation({ enabled = true }) {
  const canvasRef = useRef(null);

  useEffect(() => {
    if (!enabled) return;
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    let partikel = [], rafId;

    function ukurUlang() {
      canvas.width = canvas.offsetWidth;
      canvas.height = canvas.offsetHeight;
    }

    function buatPartikel() {
      partikel = [];
      const jumlah = (canvas.width * canvas.height) / 15000;
      for (let i = 0; i < jumlah; i++) {
        const ukuran = Math.random() * 3 + 1;
        partikel.push({
          x: Math.random() * (canvas.width - ukuran * 2 - ukuran * 2) + ukuran * 2,
          y: Math.random() * (canvas.height - ukuran * 2 - ukuran * 2) + ukuran * 2,
          vx: Math.random() * 2 - 1,
          vy: Math.random() * 2 - 1,
          ukuran,
        });
      }
    }

    function hubungkan() {
      for (let a = 0; a < partikel.length; a++) {
        for (let b = a; b < partikel.length; b++) {
          const jarak = (partikel[a].x - partikel[b].x) ** 2 + (partikel[a].y - partikel[b].y) ** 2;
          if (jarak < (canvas.width / 7) * (canvas.height / 7)) {
            const opasitas = 1 - jarak / 20000;
            ctx.strokeStyle = `rgba(29,185,160,${opasitas})`;
            ctx.lineWidth = 1;
            ctx.beginPath();
            ctx.moveTo(partikel[a].x, partikel[a].y);
            ctx.lineTo(partikel[b].x, partikel[b].y);
            ctx.stroke();
          }
        }
      }
    }

    function animasikan() {
      rafId = requestAnimationFrame(animasikan);
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      for (const p of partikel) {
        if (p.x > canvas.width || p.x < 0) p.vx = -p.vx;
        if (p.y > canvas.height || p.y < 0) p.vy = -p.vy;
        p.x += p.vx;
        p.y += p.vy;
        ctx.beginPath();
        ctx.arc(p.x, p.y, p.ukuran, 0, Math.PI * 2, false);
        ctx.fillStyle = "#1DB9A0";
        ctx.fill();
      }
      hubungkan();
    }

    function saatResize() { ukurUlang(); buatPartikel(); }
    ukurUlang();
    buatPartikel();
    animasikan();
    window.addEventListener("resize", saatResize);
    return () => {
      cancelAnimationFrame(rafId);
      window.removeEventListener("resize", saatResize);
    };
  }, [enabled]);

  const kubus = [
    { sisi: 90, top: "10%", left: "10%", putar: 26, mengambang: 8, delay: 0 },
    { sisi: 150, top: "4%", left: "80%", putar: 34, mengambang: 10, delay: 1.2 },
    { sisi: 70, top: "68%", left: "16%", putar: 22, mengambang: 7, delay: 0.6 },
    { sisi: 130, top: "74%", left: "84%", putar: 30, mengambang: 9, delay: 2 },
    { sisi: 55, top: "40%", left: "6%", putar: 18, mengambang: 6, delay: 1.6 },
  ];

  return (
    <div aria-hidden="true" style={{ position: "absolute", inset: 0, overflow: "hidden", zIndex: 1, pointerEvents: "none" }}>
      <style>{`
        @keyframes gkPutarKubus { from { transform: rotateX(0deg) rotateY(0deg); } to { transform: rotateX(360deg) rotateY(360deg); } }
        @keyframes gkMengambangKubus { from { transform: translateY(0px); } to { transform: translateY(-22px); } }
      `}</style>
      <canvas ref={canvasRef} style={{ position: "absolute", inset: 0, width: "100%", height: "100%" }} />
      {enabled && kubus.map((k, i) => (
        <div key={i} style={{
          position: "absolute", top: k.top, left: k.left, width: k.sisi, height: k.sisi,
          animation: `gkMengambangKubus ${k.mengambang}s ease-in-out ${k.delay}s infinite alternate`,
        }}>
          <div style={{ width: "100%", height: "100%", perspective: 700 }}>
            <div style={{
              width: "100%", height: "100%", position: "relative", transformStyle: "preserve-3d",
              animation: `gkPutarKubus ${k.putar}s linear infinite`,
            }}>
              {["front", "back", "right", "left", "top", "bottom"].map(sisi => {
                const setengah = k.sisi / 2;
                const transformMap = {
                  front: `translateZ(${setengah}px)`,
                  back: `rotateY(180deg) translateZ(${setengah}px)`,
                  right: `rotateY(90deg) translateZ(${setengah}px)`,
                  left: `rotateY(-90deg) translateZ(${setengah}px)`,
                  top: `rotateX(90deg) translateZ(${setengah}px)`,
                  bottom: `rotateX(-90deg) translateZ(${setengah}px)`,
                };
                return (
                  <div key={sisi} style={{
                    position: "absolute", width: "100%", height: "100%",
                    background: "linear-gradient(135deg, rgba(255,255,255,0.16), rgba(255,255,255,0.02))",
                    border: "1px solid rgba(255,255,255,0.16)", transform: transformMap[sisi],
                  }} />
                );
              })}
            </div>
          </div>
        </div>
      ))}
    </div>
  );
}
