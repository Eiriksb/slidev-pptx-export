<!-- A table that gains a column per click, and a caption that changes per click, both from $clicks. -->
<template>
  <table>
    <tbody>
      <tr>
        <th v-for="c in $clicks + 1" :key="c">Column {{ c }}</th>
      </tr>
      <tr>
        <td v-for="c in $clicks + 1" :key="c">Value {{ c }}</td>
      </tr>
    </tbody>
  </table>
  <p>Step {{ $clicks }}</p>
</template>
